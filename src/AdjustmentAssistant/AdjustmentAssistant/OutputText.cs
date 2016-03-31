using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace AdjustmentAssistant
{
    class OutputText
    {
        internal void OutputTraverse(string path, DataType.Data approximateDataType, int backCount, double angleCloseError, double coordinateCloseError, double k, List<string> col0, List<string> col1, List<string> col2, List<string> col3, List<string> col4, List<string> col5, List<string> col6, List<string> col7, List<string> col8, List<string> col9, List<string> col10)
        {
            using (StreamWriter sw = new StreamWriter(path, true))
            {
                sw.WriteLine();
                sw.WriteLine("{0,42}", "导线观测记录");
                sw.WriteLine("{0,-27}{1,-29}{2,-28}", "工程名称：", "仪器：", "天气：");
                sw.WriteLine("{0,-28}{1,-28}{2,-28}", "观测者：", "记录者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
                sw.WriteLine("----------------------------------------------------------------------------------------------------------------|");
                sw.WriteLine("|测站名 | 照准点 | 盘左读数  | 盘右读数  | 2C | 角值      | 方位角    | 平距 | 平均距离 | X         | Y         |");
                if (approximateDataType == DataType.Data.ConnectingTraverse)
                {
                    sw.WriteLine("|-------|--------|-----------|-----------|----|-----------|-----------|------|----------|-----------|-----------|");
                    for (int i = 0; i < backCount * 2; i++)
                    {
                        if (i == 0)
                        {
                            sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", col0[i], col1[i], col2[i], col3[i], col4[i], col5[i], col6[i], col7[i], col8[i], col9[i], col10[i]);
                            continue;
                        }
                        if (i == backCount)
                        {
                            sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", " ", col1[i], col2[i], col3[i], " ", " ", col6[2], col7[i / 2], " ", " ", " ");
                            continue;
                        }
                        if (i % 2 == 0 && i != 0 && i != backCount)
                        {
                            sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", " ", col1[i], col2[i], col3[i], col4[i / 2], " ", " ", col7[i / 2], " ", " ", " ");
                            continue;
                        }
                        sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", " ", col1[i], col2[i], col3[i], " ", " ", " ", " ", " ", " ", " ");
                    }
                    for (int i = 1; i < col2.Count / backCount / 2 - 1; i++)
                    {
                        sw.WriteLine("|-------|--------|-----------|-----------|----|-----------|-----------|------|----------|-----------|-----------|");
                        for (int j = 0; j < backCount * 2; j++)
                        {
                            if (j == 0)
                            {
                                sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", col0[i], col1[i * backCount * 2 + j], col2[i * backCount * 2 + j], col3[i * backCount * 2 + j], col4[(i * backCount * 2 + j) / 2], col5[i], col6[i + 2], col7[(i * backCount * 2 + j) / 2], col8[i], col9[i + 1], col10[i + 1]);
                                continue;
                            }
                            if (j % 2 == 0 && i != 0)
                            {
                                sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", " ", col1[i * backCount * 2 + j], col2[i * backCount * 2 + j], col3[i * backCount * 2 + j], col4[(i * backCount * 2 + j) / 2], " ", " ", col7[(i * backCount * 2 + j) / 2], " ", " ", " ");
                                continue;
                            }
                            sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", " ", col1[i * backCount * 2 + j], col2[i * backCount * 2 + j], col3[i * backCount * 2 + j], " ", " ", " ", " ", " ", " ", " ");
                        }
                    }
                    sw.WriteLine("|-------|--------|-----------|-----------|----|-----------|-----------|------|----------|-----------|-----------|");
                    for (int i = col2.Count - backCount * 2; i < col2.Count; i++)
                    {
                        if (i == col2.Count - backCount * 2)
                        {
                            sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", col0[i / backCount / 2], col1[i], col2[i], col3[i], col4[i / 2], col5[i / backCount / 2], col6[1], col7[i / 2], col8[i / backCount / 2], col9[1], col10[1]);
                            continue;
                        }
                        if (i % 2 == 0 && i != col2.Count - backCount * 2)
                        {
                            sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", " ", col1[i], col2[i], col3[i], col4[i / 2], " ", " ", col7[i / 2], " ", " ", " ");
                            continue;
                        }
                        sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", " ", col1[i], col2[i], col3[i], " ", " ", " ", " ", " ", " ", " ");
                    }
                    sw.WriteLine("|-------|-------------------------------------------------------------------------------------------------------|");
                    sw.WriteLine("| 备注  |类型:闭附和路线  角度闭合差:{0,-14} 坐标增量闭合差:±{1,-17} K ≈ {2,-20}|", angleCloseError, coordinateCloseError, "1/" + k);
                    sw.WriteLine("----------------------------------------------------------------------------------------------------------------|");
                }
                else if (approximateDataType == DataType.Data.OpenTraverse)
                {
                    //支导线
                    sw.WriteLine("|-------|--------|-----------|-----------|----|-----------|-----------|------|----------|-----------|-----------|");
                    for (int i = 0; i < backCount * 2; i++)
                    {
                        if (i == 0)
                        {
                            sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", col0[i], col1[i], col2[i], col3[i], col4[i], col5[i], col6[i], col7[i], col8[i], col9[i], col10[i]);
                            continue;
                        }
                        if (i == backCount)
                        {
                            sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", " ", col1[i], col2[i], col3[i], " ", " ", col6[1], col7[i / 2], " ", " ", " ");
                            continue;
                        }
                        if (i % 2 == 0 && i != 0 && i != backCount)
                        {
                            sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", " ", col1[i], col2[i], col3[i], col4[i / 2], " ", " ", col7[i / 2], " ", " ", " ");
                            continue;
                        }
                        sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", " ", col1[i], col2[i], col3[i], " ", " ", " ", " ", " ", " ", " ");
                    }
                    for (int i = 1; i < col2.Count / backCount / 2; i++)
                    {
                        sw.WriteLine("|-------|--------|-----------|-----------|----|-----------|-----------|------|----------|-----------|-----------|");
                        for (int j = 0; j < backCount * 2; j++)
                        {
                            if (j == 0)
                            {
                                sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", col0[i], col1[i * backCount * 2 + j], col2[i * backCount * 2 + j], col3[i * backCount * 2 + j], col4[(i * backCount * 2 + j) / 2], col5[i], col6[i + 1], col7[(i * backCount * 2 + j) / 2], col8[i], col9[i], col10[i]);
                                continue;
                            }
                            if (j % 2 == 0 && i != 0)
                            {
                                sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", " ", col1[i * backCount * 2 + j], col2[i * backCount * 2 + j], col3[i * backCount * 2 + j], col4[(i * backCount * 2 + j) / 2], " ", " ", col7[(i * backCount * 2 + j) / 2], " ", " ", " ");
                                continue;
                            }
                            sw.WriteLine("|{0,7}|{1,8}|{2,11}|{3,11}|{4,4}|{5,11}|{6,11}|{7,6}|{8,10}|{9,11}|{10,11}|", " ", col1[i * backCount * 2 + j], col2[i * backCount * 2 + j], col3[i * backCount * 2 + j], " ", " ", " ", " ", " ", " ", " ");
                        }
                    }
                    sw.WriteLine("|-------|-------------------------------------------------------------------------------------------------------|");
                    sw.WriteLine("| 备注  |类型:支导线路线                                                                                        |");
                    sw.WriteLine("----------------------------------------------------------------------------------------------------------------|");
                }
                sw.WriteLine("{0,-28}{1,-28}{2,-28}", "计算者：", "审核者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            }
        }

        internal void OutputLevelAngle(string path, DataType.Data approximateDataType, int backCount, double angleCloseError, List<string> col0, List<string> col1, List<string> col2, List<string> col3, List<string> col4, List<string> col5, List<string> col6)
        {
            using (StreamWriter sw = new StreamWriter(path, true))
            {
                sw.WriteLine();
                sw.WriteLine("{0,38}", "水平角观测记录");
                sw.WriteLine("{0,-25}{1,-27}{2,-26}", "工程名称：", "仪器：", "天气：");
                sw.WriteLine("{0,-26}{1,-26}{2,-26}", "观测者：", "记录者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
                sw.WriteLine("---------------------------------------------------------------------------|");
                sw.WriteLine("| 测站名 | 照准点 |  盘左读数  |  盘右读数  | 2C |    角值    |    方位角  |");
                if (approximateDataType == DataType.Data.ConnectingTraverse)
                {
                    sw.WriteLine("|--------|--------|------------|------------|----|------------|------------|");
                    for (int i = 0; i < backCount * 2; i++)
                    {
                        if (i == 0)
                        {
                            sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", col0[i], col1[i], col2[i], col3[i], col4[i], col5[i], col6[i]);
                            continue;
                        }
                        if (i == backCount)
                        {
                            sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", " ", col1[i], col2[i], col3[i], " ", " ", col6[2]);
                            continue;
                        }
                        if (i % 2 == 0 && i != 0 && i != backCount)
                        {
                            sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", " ", col1[i], col2[i], col3[i], col4[i / 2], " ", " ");
                            continue;
                        }
                        sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", " ", col1[i], col2[i], col3[i], " ", " ", " ");
                    }
                    for (int i = 1; i < col2.Count / backCount / 2 - 1; i++)
                    {
                        sw.WriteLine("|--------|--------|------------|------------|----|------------|------------|");
                        for (int j = 0; j < backCount * 2; j++)
                        {
                            if (j == 0)
                            {
                                sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", col0[i], col1[i * backCount * 2 + j], col2[i * backCount * 2 + j], col3[i * backCount * 2 + j], col4[(i * backCount * 2 + j) / 2], col5[i], col6[i + 2]);
                                continue;
                            }
                            if (j % 2 == 0 && i != 0)
                            {
                                sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", " ", col1[i * backCount * 2 + j], col2[i * backCount * 2 + j], col3[i * backCount * 2 + j], col4[(i * backCount * 2 + j) / 2], " ", " ");
                                continue;
                            }
                            sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", " ", col1[i * backCount * 2 + j], col2[i * backCount * 2 + j], col3[i * backCount * 2 + j], " ", " ", " ");
                        }
                    }
                    sw.WriteLine("|--------|--------|------------|------------|----|------------|------------|");
                    for (int i = col2.Count - backCount * 2; i < col2.Count; i++)
                    {
                        if (i == col2.Count - backCount * 2)
                        {
                            sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", col0[i / backCount / 2], col1[i], col2[i], col3[i], col4[i / 2], col5[i / backCount / 2], col6[1]);
                            continue;
                        }
                        if (i % 2 == 0 && i != col2.Count - backCount * 2)
                        {
                            sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", " ", col1[i], col2[i], col3[i], col4[i / 2], " ", " ");
                            continue;
                        }
                        sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", " ", col1[i], col2[i], col3[i], " ", " ", " ");
                    }
                    sw.WriteLine("|--------|-----------------------------------------------------------------|");
                    sw.WriteLine("|  备注  |类型:闭附合路线   角度闭合差:f = {0,-32}|", angleCloseError);
                    sw.WriteLine("---------------------------------------------------------------------------|");
                }
                else
                {
                    //支导线
                    sw.WriteLine("|--------|--------|------------|------------|----|------------|------------|");
                    for (int i = 0; i < backCount * 2; i++)
                    {
                        if (i == 0)
                        {
                            sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", col0[i], col1[i], col2[i], col3[i], col4[i], col5[i], col6[i]);
                            continue;
                        }
                        if (i == backCount)
                        {
                            sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", " ", col1[i], col2[i], col3[i], " ", " ", col6[1]);
                            continue;
                        }
                        if (i % 2 == 0 && i != 0 && i != backCount)
                        {
                            sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", " ", col1[i], col2[i], col3[i], col4[i / 2], " ", " ");
                            continue;
                        }
                        sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", " ", col1[i], col2[i], col3[i], " ", " ", " ");
                    }
                    for (int i = 1; i < col2.Count / backCount / 2; i++)
                    {
                        sw.WriteLine("|--------|--------|------------|------------|----|------------|------------|");
                        for (int j = 0; j < backCount * 2; j++)
                        {
                            if (j == 0)
                            {
                                sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", col0[i], col1[i * backCount * 2 + j], col2[i * backCount * 2 + j], col3[i * backCount * 2 + j], col4[(i * backCount * 2 + j) / 2], col5[i], col6[i + 1]);
                                continue;
                            }
                            if (j % 2 == 0 && i != 0)
                            {
                                sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", " ", col1[i * backCount * 2 + j], col2[i * backCount * 2 + j], col3[i * backCount * 2 + j], col4[(i * backCount * 2 + j) / 2], " ", " ");
                                continue;
                            }
                            sw.WriteLine("|{0,8}|{1,8}|{2,12}|{3,12}|{4,4}|{5,12}|{6,12}|", " ", col1[i * backCount * 2 + j], col2[i * backCount * 2 + j], col3[i * backCount * 2 + j], " ", " ", " ");
                        }
                    }
                    sw.WriteLine("|--------|-----------------------------------------------------------------|");
                    sw.WriteLine("|  备注  |类型:支导线路线                                                  |");
                    sw.WriteLine("---------------------------------------------------------------------------|");
                }
                sw.WriteLine("{0,-26}{1,-26}{2,-26}", "计算者：", "审核者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            }
        }

        internal void OutputPlane(string p, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5, List<string> col8, List<string> col14, List<string> col15)
        {
            using (StreamWriter sw = new StreamWriter(p, true))
            {
                sw.WriteLine();
                sw.WriteLine("{0,42}", "导线平差计算");
                sw.WriteLine("{0,-27}{1,-29}{2,-28}", "工程名称：", "仪器：", "天气：");
                sw.WriteLine("{0,-28}{1,-28}{2,-28}", "观测者：", "记录者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
                sw.WriteLine("|-----------------------------------------------------------------------------------------------------|");
                sw.WriteLine("|点名 | 近似坐标X(m)| 近似坐标Y(m)|近似边长(m)| 近似方位角|点位中误差(m)|  平差值X(m)  |  平差值Y(m)  |");
                for (int i = 0; i < col0.Count - 1; i++)
                {
                    sw.WriteLine("|-----|-------------|-------------|-----------|-----------|-------------|--------------|--------------|");
                    sw.WriteLine("|{0,5}|{1,13}|{2,13}|{3,11}|{4,11}|{5,13}|{6,14}|{7,14}|", col0[i], col1[i], col2[i], col4[i], col5[i], col8[i], col14[i], col15[i]);
                }
                sw.WriteLine("|-----|-------------|-------------|-----------|-----------|-------------|--------------|--------------|");
                sw.WriteLine("|{0,5}|{1,13}|{2,13}|{3,11}|{4,11}|{5,13}|{6,14}|{7,14}|", col0[col0.Count - 1], col1[col0.Count - 1], col2[col0.Count - 1], " ", " ", col8[col0.Count - 1], col14[col0.Count - 1], col15[col0.Count - 1]);
                sw.WriteLine("|-----------------------------------------------------------------------------------------------------|");
                sw.WriteLine("{0,-28}{1,-28}{2,-28}", "计算者：", "审核者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            }
        }

        internal void OutputPoint(string p, List<string> col0, List<string> col14, List<string> col15)
        {
            using (StreamWriter sw = new StreamWriter(p, true))
            {
                sw.WriteLine();
                sw.WriteLine("{0,28}", "控制点成果表");
                sw.WriteLine("{0,-20}{1,-20}", "工程名称：", "计算者：");
                sw.WriteLine("|--------------------------------------------------|");
                sw.WriteLine("| 点名 |     X(m)      |     Y(m)      |  高程(m)  |");
                for (int i = 0; i < col0.Count; i++)
                {
                    sw.WriteLine("|------|---------------|---------------|-----------|");
                    sw.WriteLine("|{0,6}|{1,15}|{2,15}|{3,11}|", col0[i], col14[i], col15[i], "0");
                }
                sw.WriteLine("|------|---------------|---------------|-----------|");
                sw.WriteLine("{0,-20}{1,-20}", "校核者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            }
        }

        internal void OutputResult(string p, List<string> col0, List<string> col14, List<string> col15, List<string> col18, List<string> col16, List<string> col17)
        {
            using (StreamWriter sw = new StreamWriter(p, true))
            {
                sw.WriteLine();
                sw.WriteLine("{0,65}", "平差计算成果表");
                sw.WriteLine("{0,-60}{1,-60}", "工程名称：", "计算者：");
                sw.WriteLine("|-----------------------------------------------------------------------------------------------------------------------|");
                sw.WriteLine("| 点名|    坐标X(m)    |    坐标Y(m)    | 高程(m) |角度平差值(° ' \")| 至点| 方位角(° ' \") |边长平差值(m)|高差平差值(m)|");
                for (int i = 0; i < col0.Count; i++)
                {
                    if (i == 0)
                    {
                        sw.WriteLine("|-----|----------------|----------------|---------|------------------|-----|----------------|-------------|-------------|");
                        sw.WriteLine("|{0,5}|{1,16}|{2,16}|{3,9}|{4,18}|{5,5}|{6,16}|{7,13}|{8,13}|", col0[i], col14[i], col15[i], "0", " ", col0[i + 1], col16[i], col17[i], "0");
                    }
                    else if (i == col0.Count - 1)
                    {
                        sw.WriteLine("|-----|----------------|----------------|---------|------------------|-----|----------------|-------------|-------------|");
                        sw.WriteLine("|{0,5}|{1,16}|{2,16}|{3,9}|{4,18}|{5,5}|{6,16}|{7,13}|{8,13}|", col0[i], col14[i], col15[i], "0", " ", " ", " ", " ", " ");
                    }
                    else
                    {
                        sw.WriteLine("|-----|----------------|----------------|---------|------------------|-----|----------------|-------------|-------------|");
                        sw.WriteLine("|{0,5}|{1,16}|{2,16}|{3,9}|{4,18}|{5,5}|{6,16}|{7,13}|{8,13}|", col0[i], col14[i], col15[i], "0", col18[i - 1], col0[i + 1], col16[i], col17[i], "0");
                    }
                }
                sw.WriteLine("|-----|----------------|----------------|---------|------------------|-----|----------------|-------------|-------------|");
                sw.WriteLine("{0,-60}{1,-60}", "校核者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            }
        }

        internal void OutputAccuracy(string p, double unitError, List<string> col0, List<string> col6, List<string> col7, List<string> col8, List<string> col9, List<string> col10, List<string> col11, List<string> col12, List<string> col13)
        {
            using (StreamWriter sw = new StreamWriter(p, true))
            {
                sw.WriteLine();
                sw.WriteLine("{0,65}", "精度评定表");
                sw.WriteLine("{0,-50}{1,-50}", "工程名称：", "计算者：");
                sw.WriteLine("|------------------------------------------------------------------------------------------------------------------|");
                sw.WriteLine("| 点名|  Mx(m) |  My(m) |  M(m)  |  A(m)  |  B(m)  |   F(° ' \")   |高程中误差| 至点|  方位角中误差  |边长中误差(m)|");
                for (int i = 0; i < col0.Count; i++)
                {
                    if (i == 0 || i == 1 || i == col0.Count - 2)
                    {
                        sw.WriteLine("|-----|--------|--------|--------|--------|--------|---------------|----------|-----|----------------|-------------|");
                        sw.WriteLine("|{0,5}|{1,8}|{2,8}|{3,8}|{4,8}|{5,8}|{6,15}|{7,10}|{8,5}|{9,16}|{10,13}|", col0[i], col6[i], col7[i], col8[i], " ", " ", " ", "0", col0[i + 1], col10[i], col9[i]);
                    }
                    else if (i == col0.Count - 1)
                    {
                        sw.WriteLine("|-----|--------|--------|--------|--------|--------|---------------|----------|-----|----------------|-------------|");
                        sw.WriteLine("|{0,5}|{1,8}|{2,8}|{3,8}|{4,8}|{5,8}|{6,15}|{7,10}|{8,5}|{9,16}|{10,13}|", col0[i], col6[i], col7[i], col8[i], " ", " ", " ", " ", " ", " ", " ");
                    }
                    else
                    {
                        sw.WriteLine("|-----|--------|--------|--------|--------|--------|---------------|----------|-----|----------------|-------------|");
                        sw.WriteLine("|{0,5}|{1,8}|{2,8}|{3,8}|{4,8}|{5,8}|{6,15}|{7,10}|{8,5}|{9,16}|{10,13}|", col0[i], col6[i], col7[i], col8[i], col11[i - 2], col12[i - 2], col13[i - 2], "0", col0[i + 1], col10[i], col9[i]);
                    }
                }
                sw.WriteLine("|------------------------------------------------------------------------------------------------------------------|");
                sw.WriteLine("| 备注|{0,-102}|", "单位权中误差=" + unitError);
                sw.WriteLine("|------------------------------------------------------------------------------------------------------------------|");
                sw.WriteLine("{0,-50}{1,-50}", "校核者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            }
        }

        internal void OutputBLToXY(string p, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5)
        {
            using (StreamWriter sw = new StreamWriter(p, true))
            {
                sw.WriteLine();
                sw.WriteLine("{0,40}", "高斯投影正算");
                sw.WriteLine("{0,-39}{1,-35}", "工程名称：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
                sw.WriteLine("|----------------------------------------------------------------------------|");
                sw.WriteLine("| 点名 |   纬度(° ' \")  |   经度(° ' \")  |     X(m)       |      Y(m)      |");
                for (int i = 0; i < col0.Count; i++)
                {
                    sw.WriteLine("|------|-----------------|-----------------|----------------|----------------|");
                    sw.WriteLine("|{0,6}|{1,17}|{2,17}|{3,16}|{4,16}|", col0[i], col1[i], col2[i], col4[i], col5[i]);
                }
                sw.WriteLine("|----------------------------------------------------------------------------|");
                sw.WriteLine("{0,-25}{1,-25}{2,-25}", "计算：", "复核：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            }
        }

        internal void OutputXYToBL(string p, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5)
        {
            using (StreamWriter sw = new StreamWriter(p, true))
            {
                sw.WriteLine();
                sw.WriteLine("{0,40}", "高斯投影反算");
                sw.WriteLine("{0,-39}{1,-35}", "工程名称：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
                sw.WriteLine("|----------------------------------------------------------------------------|");
                sw.WriteLine("| 点名 |     X(m)       |      Y(m)      |   纬度(° ' \")  |   经度(° ' \")  |");
                for (int i = 0; i < col0.Count; i++)
                {
                    sw.WriteLine("|------|----------------|----------------|-----------------|-----------------|");
                    sw.WriteLine("|{0,6}|{1,16}|{2,16}|{3,17}|{4,17}|", col0[i], col1[i], col2[i], col4[i], col5[i]);
                }
                sw.WriteLine("|----------------------------------------------------------------------------|");
                sw.WriteLine("{0,-25}{1,-25}{2,-25}", "计算：", "复核：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            }
        }

        internal void OutputXYToXY(string p, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5)
        {
            using (StreamWriter sw = new StreamWriter(p, true))
            {
                sw.WriteLine();
                sw.WriteLine("{0,38}", "坐标换带计算");
                sw.WriteLine("{0,-37}{1,-35}", "工程名称：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
                sw.WriteLine("|--------------------------------------------------------------------------|");
                sw.WriteLine("| 点名 |     前X(m)     |      前Y(m)    |      后X(m)    |      后Y(m)    |");
                for (int i = 0; i < col0.Count; i++)
                {
                    sw.WriteLine("|------|----------------|----------------|----------------|----------------|");
                    sw.WriteLine("|{0,6}|{1,16}|{2,16}|{3,16}|{4,16}|", col0[i], col1[i], col2[i], col4[i], col5[i]);
                }
                sw.WriteLine("|--------------------------------------------------------------------------|");
                sw.WriteLine("{0,-25}{1,-25}{2,-25}", "计算：", "复核：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            }
        }
    }
}
