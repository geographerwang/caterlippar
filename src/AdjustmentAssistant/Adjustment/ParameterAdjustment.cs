using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Adjustment
{
    public class ParameterAdjustment
    {
        private static int totalPoints;                 //总点数
        private static int knewPoints;                  //已知点数
        private static int dirGroups;                   //方向值组数
        private static int totalDir;                    //方向值总数
        private static int sideTotal;                   //边长总数
        private static int totalAzimuth;                //方位角总数
        private static double[] XY;                     //坐标数组
        private static string[] pointsNames;            //点名数组
        private static int[] dir0;                      //测站零方向在方向值数组中的位置
        private static int[] stationPointNo;            //测站点号
        private static int[] dirPointNo;                //观测方向的点号
        private static double[] dirL;                   //方向观测值
        private static double[] dirV;                   //方向观测值残差数组
        private static int[] sideStationPointNo;        //边长观测值测站点号数组
        private static int[] sideAimPointNo;            //边长观测值照准点号数组
        private static double[] sideL;                  //边长观测值数组
        private static double[] sideV;                  //边长残差数组
        private static int[] azimuthStationPointNo;     //方位角测站点点号数组
        private static int[] azimuthAimPointNo;         //方位角照准点点号数组
        private static double[] azimuthL;               //方位角观测值数组
        private static double[] azimuthV;               //方位角残差数组
        private static bool[] usable;                   //观测值可用标志数组
        private static double[] ATPA;                   //法方程系数阵
        private static double[] ATPL;                   //法方程自由项
        private static double[] dX;                     //未知参数数组（坐标改正数、定向角改正数）
        private static double dirMeanError;             //方向值中误差
        private static double azimuthMeanError;         //方位角中误差
        private static double sideError;                //边长固定误差
        private static double proportionalError;        //比例误差
        private static double pvv;                      // [pvv]
        private static double p;                        //验后单位权中误差

        public static string Calculate(string strData)
        {
            string strResult = Path.GetTempFileName();
            GetData(strData);
            GetX0Y0();
            double max = 1.0;
            while (max > 0.01)
            {
                GetATPA();
                for (int i = 0; i < knewPoints; i++)
                {
                    ATPA[GetIJ(2 * i, 2 * i)] += 1.0e20;
                    ATPA[GetIJ(2 * i + 1, 2 * i + 1)] += 1.0e20;
                }
                max = GetdX();
            }
            pvv = GetV();
            int n = totalDir + sideTotal + totalAzimuth;
            int t = 2 * (totalPoints - knewPoints) + dirGroups;
            p = Math.Sqrt(pvv / (n - t));
            using (StreamWriter sw = new StreamWriter(strResult))
            {
                sw.WriteLine("{0}", Math.Round(p, 6));
            }
            PrintResult(strResult);
            ErrorEllipse(strResult);
            return strResult;
        }

        private static void ErrorEllipse(string strResult)
        {
            double m2 = p * p;
            using (StreamWriter sw = new StreamWriter(strResult, true))
            {
                for (int i = 0; i < totalPoints; i++)
                {
                    double mx2 = ATPA[GetIJ(2 * i, 2 * i)] * m2;  //x坐标中误差的平方
                    double my2 = ATPA[GetIJ(2 * i + 1, 2 * i + 1)] * m2; //y坐标中误差的平方
                    double mxy = ATPA[GetIJ(2 * i, 2 * i + 1)] * m2; //xy坐标的协方差
                    if (Math.Sqrt(mx2 + my2) < 0.000000001)//这里的0.000000001为精度设定
                    {
                        continue;
                    }
                    double K = Math.Sqrt((mx2 - my2) * (mx2 - my2) + 4.0 * mxy * mxy);
                    double E = Math.Sqrt(0.5 * (mx2 + my2 + K));  //长轴
                    double F = Math.Sqrt(0.5 * (mx2 + my2 - K));  //短轴
                    double A; //误差椭圆长轴的方位角
                    if (Math.Abs(mxy) < 1.0e-14)  // mxy = 0
                    {
                        if (mx2 > my2)
                        {
                            A = 0.0;
                        }
                        else
                        {
                            A = 0.5 * Math.PI;
                        }
                    }
                    else   // mxy ≠ 0
                    {
                        A = Math.Atan((E * E - mx2) / mxy);
                    }
                    if (A < 0.0)
                    {
                        A += Math.PI;
                    }
                    sw.WriteLine("{0,6} {1} {2} {3}", pointsNames[i], Math.Round(E, 4), Math.Round(F, 4), A);
                }
            }
        }

        private static void PrintResult(string strResult)
        {
            int i, j, k1, k2;
            double m1, xi, yi, dxi, dyi, m2;
            using (StreamWriter sw = new StreamWriter(strResult, true))
            {
                for (i = 0; i < totalPoints; i++)
                {
                    xi = XY[2 * i];
                    yi = XY[2 * i + 1];
                    dxi = dX[2 * i];
                    dyi = dX[2 * i + 1];
                    m1 = Math.Sqrt(ATPA[GetIJ(2 * i, 2 * i)]) * p;
                    m2 = Math.Sqrt(ATPA[GetIJ(2 * i + 1, 2 * i + 1)]) * p;
                    sw.WriteLine("{0,-2} {1,-3} {2} {3} {4} {5} {6}", i + 1, pointsNames[i], Math.Round(xi, 4), Math.Round(yi, 4), Math.Round(m1, 4), Math.Round(m2, 4), Math.Round(Math.Sqrt(m1 * m1 + m2 * m2), 4));
                }
                if (dirGroups > 0)
                {
                    for (i = 0; i < dirGroups; i++)
                    {
                        k1 = stationPointNo[i];
                        for (j = dir0[i]; j < dir0[i + 1]; j++)
                        {
                            double T, S, q;
                            double[] A = new double[5];
                            int[] Ain = new int[5];
                            k2 = dirPointNo[j];
                            T = GetAB(k1, k2, A, Ain);
                            q = Math.Sqrt(GetWeightReciprocal(A, Ain)) * p;
                            S = GetCD(k1, k2, A, Ain);
                            if (j == dir0[i])
                            {
                                sw.WriteLine("{0,-3} {1} {2} {3} {4} {5} {6} {7}", pointsNames[k1], pointsNames[k2], dirL[j], dirV[j], T, q, S, Math.Sqrt(GetWeightReciprocal(A, Ain)) * p);
                            }
                            else
                            {
                                sw.WriteLine("{0,3} {1} {2} {3} {4} {5} {6} {7}", " ", pointsNames[k2], dirL[j], dirV[j], T, q, S, Math.Sqrt(GetWeightReciprocal(A, Ain)) * p);
                            }
                            if (!usable[j])
                            {
                                using (StreamWriter swErr = new StreamWriter("Error.log", true))
                                {
                                    swErr.WriteLine("\r\n" + DateTime.Now.ToString() + "：\r\n\t方向观测值有粗差");
                                }
                            }
                        }
                    }
                }

                if (sideTotal > 0)
                {
                    double[] A = new double[5];
                    int[] Ain = new int[5];
                    for (i = 0; i < sideTotal; i++)
                    {
                        int k11 = sideStationPointNo[i];
                        int k22 = sideAimPointNo[i];
                        double T = GetAB(k11, k22, A, Ain);
                        double mT = Math.Sqrt(GetWeightReciprocal(A, Ain)) * p;
                        double S = GetCD(k11, k22, A, Ain);
                        double mS = Math.Sqrt(GetWeightReciprocal(A, Ain)) * p;
                        sw.WriteLine("{0} {1} {2} {3} {4} {5} {6} {7}", pointsNames[k11], pointsNames[k22], Math.Round(sideL[i], 4), Math.Round(sideV[i], 4), T, mT, Math.Round(S, 4), Math.Round(mS, 4));
                        if (!usable[totalDir + i])
                        {
                            using (StreamWriter swErr = new StreamWriter("Error.log", true))
                            {
                                swErr.WriteLine("\r\n" + DateTime.Now.ToString() + "：\r\n\t边长观测值有粗差");
                            }
                        }
                    }
                }

                if (totalAzimuth > 0)
                {
                    double[] A = new double[5];
                    int[] Ain = new int[5];
                    for (i = 0; i < totalAzimuth; i++)
                    {
                        int k11 = azimuthStationPointNo[i];
                        int k22 = azimuthAimPointNo[i];
                        double T = GetAB(k11, k22, A, Ain);
                        double mT = Math.Sqrt(GetWeightReciprocal(A, Ain)) * p;
                        double S = GetCD(k11, k22, A, Ain);
                        double mS = Math.Sqrt(GetWeightReciprocal(A, Ain)) * p;
                        sw.WriteLine("{0} {1} {2} {3} {4} {5} {6} {7}", pointsNames[k11], pointsNames[k22], azimuthL[i], azimuthV[i], T, mT, Math.Round(S, 4), Math.Round(mS, 4));
                        if (!usable[totalDir + sideTotal + i])
                            using (StreamWriter swErr = new StreamWriter("Error.log", true))
                            {
                                swErr.WriteLine("\r\n" + DateTime.Now.ToString() + "：\r\n\t方位角观测值有粗差");
                            }
                    }
                }
            }
        }

        //权倒数的计算
        private static double GetWeightReciprocal(double[] B, int[] Bin)
        {
            int i, j, k1, k2;
            double q = 0.0;
            for (i = 0; i < 4; i++)
            {
                k1 = Bin[i];
                for (j = 0; j < 4; j++)
                {
                    k2 = Bin[j];
                    q += ATPA[GetIJ(k1, k2)] * B[i] * B[j];
                }
            }
            if (Math.Abs(q) < 1.0e-8)// 因为收舍误差q接近0时可能为负数
            {
                q = 0.0;
            }
            return q;
        }

        private static double GetV()
        {
            pvv = 0.0;
            double[] A = new double[5];
            int[] Ain = new int[5]; ;

            //  方向值残差计算
            A[4] = -1.0;

            double Pi = 1.0 / (dirMeanError * dirMeanError); //方向值的权
            for (int i = 0; i < dirGroups; i++)
            {
                int k1 = stationPointNo[i];
                Ain[4] = 2 * totalPoints + i;
                for (int j = dir0[i]; j < dir0[i + 1]; j++)
                {
                    int k2 = dirPointNo[j];
                    double T = GetAB(k1, k2, A, Ain);
                    double vj = dirV[j];
                    for (int s = 0; s < 5; s++)
                    {
                        int k = Ain[s];
                        vj += A[s] * dX[k];
                    }
                    dirV[j] = vj;
                    if (usable[j])
                    {
                        pvv += vj * vj * Pi;
                    }
                }
            }

            //  边长残差计算
            for (int i = 0; i < sideTotal; i++)
            {
                int k1 = sideStationPointNo[i];
                int k2 = sideAimPointNo[i];
                double S12 = GetCD(k1, k2, A, Ain);
                double vi = S12 - sideL[i];
                sideV[i] = vi;
                double m = sideError + proportionalError * sideL[i];
                Pi = 1.0 / (m * m);
                if (usable[totalDir + i])
                {
                    pvv += Pi * vi * vi;
                }
            }

            //  方位角残差计算
            Pi = 1.0 / (azimuthMeanError * azimuthMeanError);
            for (int i = 0; i < totalAzimuth; i++)
            {
                int k1 = azimuthStationPointNo[i];
                int k2 = azimuthAimPointNo[i];
                double T12 = GetDirection(k1, k2);
                double vi = (T12 - azimuthL[i]) * 206264.806247;
                azimuthV[i] = vi;
                if (usable[totalDir + sideTotal + i])
                {
                    pvv += Pi * vi * vi;
                }
            }
            return pvv;
        }

        private static double GetdX()
        {
            int t = 2 * totalPoints + dirGroups; //未知数个数
            if (!GetInverseMatrix(ATPA, t)) //法方程系数矩阵求逆
            {
                using (StreamWriter sw = new StreamWriter("Error.log", true))
                {
                    sw.WriteLine("\r\n" + DateTime.Now.ToString() + "\r\n\t调用ca_dX函数出错：法方程系数阵不满秩！");
                }
                Environment.Exit(0);
            }
            double max = 0.0; //坐标改正数的最大值
            for (int i = 0; i < t; i++)
            {
                double xi = 0.0;
                for (int j = 0; j < t; j++)
                {
                    xi += ATPA[GetIJ(i, j)] * ATPL[j];
                }
                dX[i] = xi;
                if (i < 2 * totalPoints)
                {
                    XY[i] += xi;
                    if (Math.Abs(xi) > max)
                    {
                        max = Math.Abs(xi);
                    }
                }
            }
            return max;
        }

        private static bool GetInverseMatrix(double[] a, int n)
        {
            double[] a0 = new double[n];
            for (int k = 0; k < n; k++)
            {
                double a00 = a[0];
                if (a00 + 1.0 == 1.0)
                {
                    return false;
                }
                for (int i = 1; i < n; i++)
                {
                    double ai0 = a[i * (i + 1) / 2];
                    if (i <= n - k - 1)
                    {
                        a0[i] = -ai0 / a00;
                    }
                    else
                    {
                        a0[i] = ai0 / a00;
                    }
                    for (int j = 1; j <= i; j++)
                    {
                        a[(i - 1) * i / 2 + j - 1] = a[i * (i + 1) / 2 + j] + ai0 * a0[j];
                    }
                }
                for (int i = 1; i < n; i++)
                {
                    a[(n - 1) * n / 2 + i - 1] = a0[i];
                }
                a[n * (n + 1) / 2 - 1] = 1.0 / a00;
            }
            return true;
        }

        private static void GetATPA()
        {
            const double ROU = 2.062648062470963552e+05;
            double[] B = new double[5];
            int[] Bin = new int[5];
            int t = 2 * totalPoints + dirGroups;
            int tt = t * (t + 1) / 2;
            for (int i = 0; i <= tt - 1; i++)
            {
                ATPA[i] = 0.0;
            }
            for (int i = 0; i <= t - 1; i++)
            {
                ATPL[i] = 0.0;
            }

            //  方向值组成法方程
            double Pi = 1.0 / (dirMeanError * dirMeanError);
            B[4] = -1.0;
            for (int i = 0; i < dirGroups; i++)
            {
                int k1 = stationPointNo[i];
                Bin[4] = 2 * totalPoints + i; //定向角改正数的未知数编号
                double z = 0;     //定向角近似值
                for (int j = dir0[i]; j < dir0[i + 1]; j++)
                {
                    if (!usable[j])
                    {
                        continue;
                    }
                    int k2 = dirPointNo[j];
                    double T = GetDirection(k1, k2); //返回值：方位角
                    z = T - dirL[j];
                    if (z < 0.0)
                    {
                        z += 2.0 * Math.PI;
                    }
                    break;
                }
                for (int j = dir0[i]; j < dir0[i + 1]; j++)
                {
                    int k2 = dirPointNo[j];
                    double T12 = GetAB(k1, k2, B, Bin);
                    double Lj = T12 - dirL[j];
                    if (Lj < 0.0)
                    {
                        Lj += 2.0 * Math.PI;
                    }
                    Lj = (Lj - z) * ROU;
                    dirV[j] = Lj; //自由项放在V数组里，计算残差时使用
                    if (usable[j])
                    {
                        GetATPAi(B, Bin, Pi, Lj, 5);
                    }
                }
            }

            //  边长组成法方程
            for (int i = 0; i < sideTotal; i++)
            {
                if (!usable[i + totalDir])
                {
                    continue;
                }
                int k1 = sideStationPointNo[i];
                int k2 = sideAimPointNo[i];
                double m = sideError + proportionalError * sideL[i];
                double pi = 1.0 / (m * m + 1.0e-15); // 边长的权
                double S12 = GetCD(k1, k2, B, Bin);
                double Li = S12 - sideL[i];
                GetATPAi(B, Bin, pi, Li, 4);
            }

            //  方位角组成法方程
            Pi = 1.0 / (azimuthMeanError * azimuthMeanError + 1.0e-15); //方位角的权
            for (int i = 0; i < totalAzimuth; i++)
            {
                if (!usable[i + totalDir + sideTotal])
                {
                    continue;
                }
                int k1 = azimuthStationPointNo[i];
                int k2 = azimuthAimPointNo[i];
                double T12 = GetAB(k1, k2, B, Bin);
                double Li = (T12 - azimuthL[i]) * ROU;
                GetATPAi(B, Bin, Pi, Li, 4);
            }
        }

        private static double GetCD(int k1, int k2, double[] A, int[] Ain)
        {
            double dx = XY[2 * k2] - XY[2 * k1];
            double dy = XY[2 * k2 + 1] - XY[2 * k1 + 1];
            double s = Math.Sqrt(dx * dx + dy * dy);
            A[0] = -dx / s;
            Ain[0] = 2 * k1;
            A[1] = -dy / s;
            Ain[1] = 2 * k1 + 1;
            A[2] = dx / s;
            Ain[2] = 2 * k2;
            A[3] = dy / s;
            Ain[3] = 2 * k2 + 1;
            return s;
        }

        private static void GetATPAi(double[] B, int[] Bin, double p, double Li, int m)
        {
            int k, s, i, j;
            double ai, aj;
            for (k = 0; k < m; k++)
            {
                i = Bin[k];
                ai = B[k];
                ATPL[i] -= p * ai * Li;
                for (s = 0; s < m; s++)
                {
                    j = Bin[s];
                    if (i > j)
                    {
                        continue;
                    }
                    aj = B[s];
                    ATPA[GetIJ(i, j)] += p * ai * aj;
                }
            }
        }

        //矩阵下边的获取
        private static int GetIJ(int i, int j)
        {
            return (i >= j) ? i * (i + 1) / 2 + j : j * (j + 1) / 2 + i;
        }

        private static double GetAB(int k1, int k2, double[] A, int[] Ain)
        {
            const double ROU = 2.062648062470963552e+05;
            double dx = XY[2 * k2] - XY[2 * k1];
            double dy = XY[2 * k2 + 1] - XY[2 * k1 + 1];
            double s2 = dx * dx + dy * dy;
            A[0] = dy / s2 * ROU;
            Ain[0] = 2 * k1;
            A[1] = -dx / s2 * ROU;
            Ain[1] = 2 * k1 + 1;
            A[2] = -dy / s2 * ROU;
            Ain[2] = 2 * k2;
            A[3] = dx / s2 * ROU;
            Ain[3] = 2 * k2 + 1;
            double T = Math.Atan2(dy, dx);
            if (T < 0.0)
            {
                T = T + 2.0 * Math.PI;
            }
            return T;
        }

        private static void GetX0Y0()
        {
            int unknow = totalPoints - knewPoints; //未知点数
            //设置未知点标志,未知点的点号从m_knPnumber开始
            for (int i = knewPoints; i < totalPoints; i++)
            {
                XY[2 * i] = 1.0e30;
            }
            for (int No = 1; ; No++)
            {
                if (unknow == 0)
                {
                    return;
                }
                if (No > (totalPoints - knewPoints))
                {
                    using (StreamWriter sw = new StreamWriter("Error.log", true))
                    {
                        sw.WriteLine("\r\n" + DateTime.Now.ToString() + ": \r\n\t部分点计算不出近似坐标:");
                        for (int k = 0; k < totalPoints; k++)
                        {
                            if (XY[2 * k] > 1.0e29)
                            {
                                sw.WriteLine("{0}", pointsNames[k]);
                            }
                        }
                    }
                    Environment.Exit(0);
                }
                for (int i = 0; i < dirGroups; i++) // 按方向组循环，遍历方向观测值
                {
                    int k1 = stationPointNo[i];  //测站点号
                    double x1 = XY[2 * k1];   //测站点的x坐标
                    double y1 = XY[2 * k1 + 1]; //测站点的y坐标
                    if (x1 > 1.0e29)//测站点是未知点，转下一方向组
                    {
                        continue;
                    }
                    int j0 = dir0[i]; //本方向组首方向观测值的序号
                    for (int j = j0; j < dir0[i + 1]; j++)
                    {
                        int k2 = dirPointNo[j]; //照准点号
                        double T12;
                        if (XY[2 * k1] < 1.0e29 && XY[2 * k2] < 1.0e29) //k1、k2都是已知点
                        {
                            T12 = GetDirection(k1, k2); //用坐标计算方位角
                        }
                        else//在方位角数组中查找起始方位角
                        {
                            T12 = GetAzimuth(k1, k2);
                        }
                        if (T12 > 1.0e29)//无起始方位角
                        {
                            continue;
                        }
                        for (int k = j0; k < dir0[i + 1]; k++)
                        {
                            int k3 = dirPointNo[k];
                            double x3 = XY[2 * k3];
                            if (x3 < 1.0e29)//k3是已知点
                            {
                                continue;
                            }
                            double S13 = GetSide(k1, k3); //在边长数组中查找边长
                            if (S13 > 1.0e29)//无边长观测值
                            {
                                continue;
                            }
                            double T13 = T12 + dirL[k] - dirL[j];
                            x3 = S13 * Math.Cos(T13);
                            double y3 = S13 * Math.Sin(T13);
                            XY[2 * k3] = x1 + x3;
                            XY[2 * k3 + 1] = y1 + y3;
                            unknow--;
                        }
                    }
                }
            }
        }

        //查找边长观测值
        private static double GetSide(int k1, int k2)
        {
            for (int i = 0; i < sideTotal; i++)
            {
                if (k1 == sideStationPointNo[i] && k2 == sideAimPointNo[i])
                {
                    return sideL[i];
                }
                if (k1 == sideAimPointNo[i] && k2 == sideStationPointNo[i])
                {
                    return sideL[i];
                }
            }
            return 1.0e30;
        }

        //在方位角观测值中查找
        private static double GetAzimuth(int k1, int k2)
        {
            for (int i = 0; i < totalAzimuth; i++)
            {
                if (k1 == azimuthStationPointNo[i] && k2 == azimuthAimPointNo[i])
                {
                    return azimuthL[i];
                }
                if (k1 == azimuthAimPointNo[i] && k2 == azimuthStationPointNo[i])
                {
                    double T = azimuthL[i] + Math.PI;
                    if (T > 2.0 * Math.PI)
                    {
                        T = T - 2.0 * Math.PI;
                    }
                    return T;
                }
            }
            return 1.0e30;
        }

        //平面方位角计算
        private static double GetDirection(int k1, int k2)
        {
            double dx = XY[2 * k2] - XY[2 * k1];
            double dy = XY[2 * k2 + 1] - XY[2 * k1 + 1];
            double T = Math.Atan2(dy, dx);
            if (T < 0.0)
            {
                T = T + 2.0 * Math.PI;
            }
            return T;
        }

        //<测试>该模块是否正确的读取数据
        private static void GetData(string strData)
        {
            using (StreamReader sr = new StreamReader(strData))
            {
                string[] strSplit = { " " };
                string strNumber = sr.ReadLine();
                string[] strArrNumber = strNumber.Split(strSplit, StringSplitOptions.RemoveEmptyEntries);
                totalPoints = Convert.ToInt32(strArrNumber[0]);
                knewPoints = Convert.ToInt32(strArrNumber[1]);
                dirGroups = Convert.ToInt32(strArrNumber[2]);
                totalDir = Convert.ToInt32(strArrNumber[3]);
                sideTotal = Convert.ToInt32(strArrNumber[4]);
                totalAzimuth = Convert.ToInt32(strArrNumber[5]);
                XY = new double[2 * totalPoints];
                pointsNames = new string[totalPoints];
                if (dirGroups > 0)
                {
                    dir0 = new int[dirGroups + 1]; //各组首方向观测值的序号
                    stationPointNo = new int[dirGroups];   //测站点号
                    dirPointNo = new int[totalDir];   //照准点号
                    dirL = new double[totalDir];   //方向值
                    dirV = new double[totalDir];   //粗差，V之前，放自由项l
                }
                if (sideTotal > 0)//为边长观测值数组申请内存
                {
                    sideStationPointNo = new int[sideTotal];  //测站点号
                    sideAimPointNo = new int[sideTotal];  //照准点号
                    sideL = new double[sideTotal];  //边长观测值
                    sideV = new double[sideTotal];  //边长残差
                }
                if (totalAzimuth > 0)//为方位角观测值数组申请内存
                {
                    azimuthStationPointNo = new int[sideTotal];  //测站点号
                    azimuthAimPointNo = new int[sideTotal];  //照准点号
                    azimuthL = new double[sideTotal];  //方位角观测值
                    azimuthV = new double[sideTotal];  //方位角残差
                }
                // 观测值可用标志数组
                int n = totalDir + sideTotal + totalAzimuth;
                usable = new bool[n];
                for (int i = 0; i < n; i++)
                {
                    usable[i] = true;
                }
                int t = 2 * totalPoints + dirGroups; // 未知参数总数
                int tt = t * (t + 1) / 2;
                ATPL = new double[t];   //法方程自由项
                ATPA = new double[tt];  //系数矩阵
                dX = new double[t];     //未知数向量
                int unPnumber = totalPoints - knewPoints;
                string strError = sr.ReadLine();
                string[] strArrError = strError.Split(strSplit, StringSplitOptions.RemoveEmptyEntries);
                dirMeanError = Convert.ToDouble(strArrError[0]);
                sideError = Convert.ToDouble(strArrError[1]);
                proportionalError = Convert.ToDouble(strArrError[2]);
                azimuthMeanError = Convert.ToDouble(strArrError[3]);
                for (int i = 0; i < knewPoints; i++)
                {
                    string strPKnew = sr.ReadLine();
                    string[] strArrpKnew = strPKnew.Split(strSplit, StringSplitOptions.RemoveEmptyEntries);
                    int k = GetStationNumber(strArrpKnew[0]);
                    XY[2 * k] = Convert.ToDouble(strArrpKnew[1]);
                    XY[2 * k + 1] = Convert.ToDouble(strArrpKnew[2]);
                }
                if (dirGroups > 0)
                {
                    dir0[0] = 0;
                    for (int i = 0; i < dirGroups; i++)
                    {
                        int ni; // ni: 测站方向数
                        string strDir = sr.ReadLine();
                        string[] strArrDir = strDir.Split(strSplit, StringSplitOptions.RemoveEmptyEntries);
                        ni = Convert.ToInt32(strArrDir[1]);
                        stationPointNo[i] = GetStationNumber(strArrDir[0]);
                        dir0[i + 1] = dir0[i] + ni;
                        for (int j = dir0[i]; j < dir0[i + 1]; j++)
                        {
                            string strDirVal = sr.ReadLine();
                            string[] strArrDirVal = strDirVal.Split(strSplit, StringSplitOptions.RemoveEmptyEntries);
                            dirPointNo[j] = GetStationNumber(strArrDirVal[0]); //照准点号
                            dirL[j] = Convert.ToDouble(strArrDirVal[1]);
                        }
                    }
                }
                for (int i = 0; i < sideTotal; i++)  //读取边长
                {
                    string strSide = sr.ReadLine();
                    string[] strArrSide = strSide.Split(strSplit, StringSplitOptions.RemoveEmptyEntries);
                    sideStationPointNo[i] = GetStationNumber(strArrSide[0]);
                    sideAimPointNo[i] = GetStationNumber(strArrSide[1]);
                    sideL[i] = Convert.ToDouble(strArrSide[2]);
                }
                for (int i = 0; i < totalAzimuth; i++)  //读取方位角
                {
                    string strDir = sr.ReadLine();
                    string[] strArrDir = strDir.Split(strSplit, StringSplitOptions.RemoveEmptyEntries);
                    azimuthStationPointNo[i] = GetStationNumber(strArrDir[0]);
                    azimuthAimPointNo[i] = GetStationNumber(strArrDir[1]);
                    azimuthL[i] = Convert.ToDouble(strArrDir[2]);
                }
            }
        }

        private static int GetStationNumber(string name)
        {
            int i;
            for (i = 0; i < totalPoints; i++)
            {
                if (pointsNames[i] == null)
                {
                    break;
                }
                if (name == pointsNames[i])
                {
                    return i;
                }
            }
            if (i < totalPoints)//已经编过点号的点数小于总点数
            {
                pointsNames[i] = name;
                return i;
            }
            else
            {
                return -1;
            }
        }
    }
}
