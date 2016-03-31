using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace AdjustmentAssistant
{
    class OutputExcel
    {
        internal void OutputTraverse(string path, DataType.Data approximateDataType, int backCount, double angleCloseError, double coordinateCloseError, double k, List<string> col0, List<string> col1, List<string> col2, List<string> col3, List<string> col4, List<string> col5, List<string> col6, List<string> col7, List<string> col8, List<string> col9, List<string> col10)
        {
            object nothing = Missing.Value;
            object fileFormat = XlFileFormat.xlExcel8;
            object fileName = path;
            Microsoft.Office.Interop.Excel._Application excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Workbook workBook = excelApp.Workbooks.Add(true);
            Worksheet worksheet = workBook.ActiveSheet as Worksheet;
            Range titleRange = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, 11]);
            titleRange.MergeCells = true;
            titleRange.Value2 = "导线观测记录";
            titleRange.Font.Name = "黑体";
            titleRange.Font.Size = 12;
            titleRange.Font.Bold = true;
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            Range projectNameRange = worksheet.get_Range(worksheet.Cells[2, 1], worksheet.Cells[2, 3]);
            projectNameRange.MergeCells = true;
            projectNameRange.Value2 = "工程名称：";
            Range instrumentRange = worksheet.get_Range(worksheet.Cells[2, 5], worksheet.Cells[2, 7]);
            instrumentRange.MergeCells = true;
            instrumentRange.Value2 = "仪器：";
            Range weatherRange = worksheet.get_Range(worksheet.Cells[2, 9], worksheet.Cells[2, 11]);
            weatherRange.MergeCells = true;
            weatherRange.Value2 = "天气：";
            Range observerRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[3, 3]);
            observerRange.MergeCells = true;
            observerRange.Value2 = "观测者：";
            Range recorderRange = worksheet.get_Range(worksheet.Cells[3, 5], worksheet.Cells[3, 7]);
            recorderRange.MergeCells = true;
            recorderRange.Value2 = "记录者：";
            Range dateRange = worksheet.get_Range(worksheet.Cells[3, 9], worksheet.Cells[3, 11]);
            dateRange.MergeCells = true;
            dateRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            Range calculateRange = worksheet.get_Range(worksheet.Cells[col1.Count + 7, 1], worksheet.Cells[col1.Count + 7, 3]);
            calculateRange.MergeCells = true;
            calculateRange.Value2 = "计算者：";
            Range assessmentRange = worksheet.get_Range(worksheet.Cells[col1.Count + 7, 5], worksheet.Cells[col1.Count + 7, 7]);
            assessmentRange.MergeCells = true;
            assessmentRange.Value2 = "审核者：";
            Range endDateRange = worksheet.get_Range(worksheet.Cells[col1.Count + 7, 9], worksheet.Cells[col1.Count + 7, 11]);
            endDateRange.MergeCells = true;
            endDateRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");

            Range borderRange = worksheet.get_Range(worksheet.Cells[4, 1], worksheet.Cells[col1.Count + 6, 11]);
            borderRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, null, null);
            borderRange.Borders.Weight = XlBorderWeight.xlThin;
            Range stationRange = worksheet.get_Range(worksheet.Cells[4, 1], worksheet.Cells[5, 1]);
            stationRange.MergeCells = true;
            stationRange.Value2 = "测站名";
            Range pointRange = worksheet.get_Range(worksheet.Cells[4, 2], worksheet.Cells[5, 2]);
            pointRange.MergeCells = true;
            pointRange.Value2 = "照准点";
            Range levelAngleRange = worksheet.get_Range(worksheet.Cells[4, 3], worksheet.Cells[4, 7]);
            levelAngleRange.MergeCells = true;
            levelAngleRange.Value2 = "水平角(° ' \")";
            worksheet.Cells[5, 3] = "盘左";
            worksheet.Cells[5, 4] = "盘右";
            worksheet.Cells[5, 5] = "2C(\")";
            worksheet.Cells[5, 6] = "角值";
            worksheet.Cells[5, 7] = "方位角";
            Range distanceRange = worksheet.get_Range(worksheet.Cells[4, 8], worksheet.Cells[4, 9]);
            distanceRange.MergeCells = true;
            distanceRange.Value2 = "距离";
            worksheet.Cells[5, 8] = "平距";
            worksheet.Cells[5, 9] = "平均距离";
            Range coordinateRange = worksheet.get_Range(worksheet.Cells[4, 10], worksheet.Cells[4, 11]);
            coordinateRange.MergeCells = true;
            coordinateRange.Value2 = "坐标值(m)";
            worksheet.Cells[5, 10] = "X";
            worksheet.Cells[5, 11] = "Y";

            if (approximateDataType == DataType.Data.ConnectingTraverse)
            {
                for (int i = 0; i < backCount * 2; i++)
                {
                    if (i == 0)
                    {
                        Range columnRange1 = worksheet.get_Range(worksheet.Cells[i + 6, 1], worksheet.Cells[backCount * 2 + 5, 1]);
                        columnRange1.MergeCells = true;
                        columnRange1.Value2 = col0[i];
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                        columnRange5.MergeCells = true;
                        columnRange5.Value2 = col4[i];
                        Range columnRange6 = worksheet.get_Range(worksheet.Cells[i + 6, 6], worksheet.Cells[backCount * 2 + 5, 6]);
                        columnRange6.MergeCells = true;
                        columnRange6.Value2 = col5[i];
                        Range columnRange7 = worksheet.get_Range(worksheet.Cells[i + 6, 7], worksheet.Cells[backCount + 5, 7]);
                        columnRange7.MergeCells = true;
                        columnRange7.Value2 = col6[i];
                        Range columnRange8 = worksheet.get_Range(worksheet.Cells[i + 6, 8], worksheet.Cells[i + 7, 8]);
                        columnRange8.MergeCells = true;
                        columnRange8.Value2 = col7[i];
                        Range columnRange9 = worksheet.get_Range(worksheet.Cells[i + 6, 9], worksheet.Cells[backCount * 2 + 5, 9]);
                        columnRange9.MergeCells = true;
                        columnRange9.Value2 = col8[i];
                        Range columnRange10 = worksheet.get_Range(worksheet.Cells[i + 6, 10], worksheet.Cells[backCount * 2 + 5, 10]);
                        columnRange10.MergeCells = true;
                        columnRange10.Value2 = col9[i];
                        Range columnRange11 = worksheet.get_Range(worksheet.Cells[i + 6, 11], worksheet.Cells[backCount * 2 + 5, 11]);
                        columnRange11.MergeCells = true;
                        columnRange11.Value2 = col10[i];
                        continue;
                    }
                    if (i == backCount)
                    {
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        if (backCount > 1)
                        {
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[i / 2];
                        }
                        Range columnRange7 = worksheet.get_Range(worksheet.Cells[i + 6, 7], worksheet.Cells[backCount * 2 + 5, 7]);
                        columnRange7.MergeCells = true;
                        columnRange7.Value2 = col6[2];
                        if (backCount > 1)
                        {
                            Range columnRange8 = worksheet.get_Range(worksheet.Cells[i + 6, 8], worksheet.Cells[i + 7, 8]);
                            columnRange8.MergeCells = true;
                            columnRange8.Value2 = col7[i / 2];
                        }
                        continue;
                    }
                    if (i % 2 == 0 && i != 0 && i != backCount)
                    {
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        if (backCount > 1)
                        {
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[i / 2];
                        }
                        if (backCount > 1)
                        {
                            Range columnRange8 = worksheet.get_Range(worksheet.Cells[i + 6, 8], worksheet.Cells[i + 7, 8]);
                            columnRange8.MergeCells = true;
                            columnRange8.Value2 = col7[i / 2];
                        }
                        continue;
                    }
                    worksheet.Cells[i + 6, 2] = col1[i];
                    worksheet.Cells[i + 6, 3] = col2[i];
                    worksheet.Cells[i + 6, 4] = col3[i];
                }
                for (int i = 1; i < col1.Count / backCount / 2 - 1; i++)
                {
                    for (int j = 0; j < backCount * 2; j++)
                    {
                        if (j == 0)
                        {
                            Range columnRange1 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 1], worksheet.Cells[(i + 1) * backCount * 2 + 5, 1]);
                            columnRange1.MergeCells = true;
                            columnRange1.Value2 = col0[i];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 2] = col1[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 3] = col2[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 4] = col3[i * backCount * 2 + j];
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 5], worksheet.Cells[i * backCount * 2 + 7 + j, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[(i * backCount * 2 + j) / 2];
                            Range columnRange6 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 6], worksheet.Cells[(i + 1) * backCount * 2 + 5, 6]);
                            columnRange6.MergeCells = true;
                            columnRange6.Value2 = col5[i];
                            Range columnRange7 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 7], worksheet.Cells[(i + 1) * backCount * 2 + 5, 7]);
                            columnRange7.MergeCells = true;
                            columnRange7.Value2 = col6[i + 2];
                            Range columnRange8 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 8], worksheet.Cells[i * backCount * 2 + 7 + j, 8]);
                            columnRange8.MergeCells = true;
                            columnRange8.Value2 = col7[(i * backCount * 2 + j) / 2];
                            Range columnRange9 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 9], worksheet.Cells[(i + 1) * backCount * 2 + 5, 9]);
                            columnRange9.MergeCells = true;
                            columnRange9.Value2 = col8[i];
                            Range columnRange10 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 10], worksheet.Cells[(i + 1) * backCount * 2 + 5, 10]);
                            columnRange10.MergeCells = true;
                            columnRange10.Value2 = col9[i + 1];
                            Range columnRange11 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 11], worksheet.Cells[(i + 1) * backCount * 2 + 5, 11]);
                            columnRange11.MergeCells = true;
                            columnRange11.Value2 = col10[i + 1];
                            continue;
                        }
                        if (j % 2 == 0 && j != 0)
                        {
                            worksheet.Cells[i * backCount * 2 + 6 + j, 2] = col1[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 3] = col2[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 4] = col3[i * backCount * 2 + j];
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 5], worksheet.Cells[i * backCount * 2 + 7 + j, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[(i * backCount * 2 + j) / 2];
                            Range columnRange8 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 8], worksheet.Cells[i * backCount * 2 + 7 + j, 8]);
                            columnRange8.MergeCells = true;
                            columnRange8.Value2 = col7[(i * backCount * 2 + j) / 2];
                            continue;
                        }
                        worksheet.Cells[i * backCount * 2 + 6 + j, 2] = col1[i * backCount * 2 + j];
                        worksheet.Cells[i * backCount * 2 + 6 + j, 3] = col2[i * backCount * 2 + j];
                        worksheet.Cells[i * backCount * 2 + 6 + j, 4] = col3[i * backCount * 2 + j];
                    }
                }
                for (int i = col1.Count - backCount * 2; i < col1.Count; i++)
                {
                    if (i == col1.Count - backCount * 2)
                    {
                        Range columnRange1 = worksheet.get_Range(worksheet.Cells[i + 6, 1], worksheet.Cells[i + backCount * 2 + 5, 1]);
                        columnRange1.MergeCells = true;
                        columnRange1.Value2 = col0[i / backCount / 2];
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                        columnRange5.MergeCells = true;
                        columnRange5.Value2 = col4[i / 2];
                        Range columnRange6 = worksheet.get_Range(worksheet.Cells[i + 6, 6], worksheet.Cells[i + backCount * 2 + 5, 6]);
                        columnRange6.MergeCells = true;
                        columnRange6.Value2 = col5[i / backCount / 2];
                        Range columnRange7 = worksheet.get_Range(worksheet.Cells[i + 6, 7], worksheet.Cells[i + backCount * 2 + 5, 7]);
                        columnRange7.MergeCells = true;
                        columnRange7.Value2 = col6[1];
                        Range columnRange8 = worksheet.get_Range(worksheet.Cells[i + 6, 8], worksheet.Cells[i + 7, 8]);
                        columnRange8.MergeCells = true;
                        columnRange8.Value2 = col7[i / 2];
                        Range columnRange9 = worksheet.get_Range(worksheet.Cells[i + 6, 9], worksheet.Cells[i + backCount * 2 + 5, 9]);
                        columnRange9.MergeCells = true;
                        columnRange9.Value2 = col8[i / backCount / 2];
                        Range columnRange10 = worksheet.get_Range(worksheet.Cells[i + 6, 10], worksheet.Cells[i + backCount * 2 + 5, 10]);
                        columnRange10.MergeCells = true;
                        columnRange10.Value2 = col9[1];
                        Range columnRange11 = worksheet.get_Range(worksheet.Cells[i + 6, 11], worksheet.Cells[i + backCount * 2 + 5, 11]);
                        columnRange11.MergeCells = true;
                        columnRange11.Value2 = col10[1];
                        continue;
                    }
                    if (i % 2 == 0 && i != col1.Count - backCount * 2)
                    {
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                        columnRange5.MergeCells = true;
                        columnRange5.Value2 = col4[i / 2];
                        Range columnRange8 = worksheet.get_Range(worksheet.Cells[i + 6, 8], worksheet.Cells[i + 7, 8]);
                        columnRange8.MergeCells = true;
                        columnRange8.Value2 = col7[i / 2];
                        continue;
                    }
                    worksheet.Cells[i + 6, 2] = col1[i];
                    worksheet.Cells[i + 6, 3] = col2[i];
                    worksheet.Cells[i + 6, 4] = col3[i];
                }
                worksheet.Cells[col1.Count + 6, 1] = "备注";
                Range rowRemarkRange = worksheet.get_Range(worksheet.Cells[col1.Count + 6, 2], worksheet.Cells[col1.Count + 6, 11]);
                rowRemarkRange.MergeCells = true;
                rowRemarkRange.Value2 = "类型:闭符合路线   角度闭合差:" + angleCloseError + " 坐标增量闭合差:±" + coordinateCloseError + " K≈1/" + k;
            }
            else if (approximateDataType == DataType.Data.OpenTraverse)
            {
                for (int i = 0; i < backCount * 2; i++)
                {
                    if (i == 0)
                    {
                        Range columnRange1 = worksheet.get_Range(worksheet.Cells[i + 6, 1], worksheet.Cells[backCount * 2 + 5, 1]);
                        columnRange1.MergeCells = true;
                        columnRange1.Value2 = col0[i];
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                        columnRange5.MergeCells = true;
                        columnRange5.Value2 = col4[i];
                        Range columnRange6 = worksheet.get_Range(worksheet.Cells[i + 6, 6], worksheet.Cells[backCount * 2 + 5, 6]);
                        columnRange6.MergeCells = true;
                        columnRange6.Value2 = col5[i];
                        Range columnRange7 = worksheet.get_Range(worksheet.Cells[i + 6, 7], worksheet.Cells[backCount + 5, 7]);
                        columnRange7.MergeCells = true;
                        columnRange7.Value2 = col6[i];
                        Range columnRange8 = worksheet.get_Range(worksheet.Cells[i + 6, 8], worksheet.Cells[i + 7, 8]);
                        columnRange8.MergeCells = true;
                        columnRange8.Value2 = col7[i];
                        Range columnRange9 = worksheet.get_Range(worksheet.Cells[i + 6, 9], worksheet.Cells[backCount * 2 + 5, 9]);
                        columnRange9.MergeCells = true;
                        columnRange9.Value2 = col8[i];
                        Range columnRange10 = worksheet.get_Range(worksheet.Cells[i + 6, 10], worksheet.Cells[backCount * 2 + 5, 10]);
                        columnRange10.MergeCells = true;
                        columnRange10.Value2 = col9[i];
                        Range columnRange11 = worksheet.get_Range(worksheet.Cells[i + 6, 11], worksheet.Cells[backCount * 2 + 5, 11]);
                        columnRange11.MergeCells = true;
                        columnRange11.Value2 = col10[i];
                        continue;
                    }
                    if (i == backCount)
                    {
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        if (backCount > 1)
                        {
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[i / 2];
                        }
                        Range columnRange7 = worksheet.get_Range(worksheet.Cells[i + 6, 7], worksheet.Cells[backCount + 5, 7]);
                        columnRange7.MergeCells = true;
                        columnRange7.Value2 = col6[1];
                        if (backCount > 1)
                        {
                            Range columnRange8 = worksheet.get_Range(worksheet.Cells[i + 6, 8], worksheet.Cells[i + 7, 8]);
                            columnRange8.MergeCells = true;
                            columnRange8.Value2 = col7[i / 2];
                        }
                        continue;
                    }
                    if (i % 2 == 0 && i != 0 && i != backCount)
                    {
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        if (backCount > 1)
                        {
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[i / 2];
                        }
                        if (backCount > 1)
                        {
                            Range columnRange8 = worksheet.get_Range(worksheet.Cells[i + 6, 8], worksheet.Cells[i + 7, 8]);
                            columnRange8.MergeCells = true;
                            columnRange8.Value2 = col7[i / 2];
                        }
                        continue;
                    }
                    worksheet.Cells[i + 6, 2] = col1[i];
                    worksheet.Cells[i + 6, 3] = col2[i];
                    worksheet.Cells[i + 6, 4] = col3[i];
                }
                for (int i = 1; i < col1.Count / backCount / 2; i++)
                {
                    for (int j = 0; j < backCount * 2; j++)
                    {
                        if (j == 0)
                        {
                            Range columnRange1 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 1], worksheet.Cells[(i + 1) * backCount * 2 + 5, 1]);
                            columnRange1.MergeCells = true;
                            columnRange1.Value2 = col0[i];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 2] = col1[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 3] = col2[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 4] = col3[i * backCount * 2 + j];
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 5], worksheet.Cells[i * backCount * 2 + 7 + j, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[(i * backCount * 2 + j) / 2];
                            Range columnRange6 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 6], worksheet.Cells[(i + 1) * backCount * 2 + 5, 6]);
                            columnRange6.MergeCells = true;
                            columnRange6.Value2 = col5[i];
                            Range columnRange7 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 7], worksheet.Cells[(i + 1) * backCount * 2 + 5, 7]);
                            columnRange7.MergeCells = true;
                            columnRange7.Value2 = col6[i + 1];
                            Range columnRange8 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 8], worksheet.Cells[i * backCount * 2 + 7 + j, 8]);
                            columnRange8.MergeCells = true;
                            columnRange8.Value2 = col7[(i * backCount * 2 + j) / 2];
                            Range columnRange9 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 9], worksheet.Cells[(i + 1) * backCount * 2 + 5, 9]);
                            columnRange9.MergeCells = true;
                            columnRange9.Value2 = col8[i];
                            Range columnRange10 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 10], worksheet.Cells[(i + 1) * backCount * 2 + 5, 10]);
                            columnRange10.MergeCells = true;
                            columnRange10.Value2 = col9[i];
                            Range columnRange11 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 11], worksheet.Cells[(i + 1) * backCount * 2 + 5, 11]);
                            columnRange11.MergeCells = true;
                            columnRange11.Value2 = col10[i];
                            continue;
                        }
                        if (j % 2 == 0 && j != 0)
                        {
                            worksheet.Cells[i * backCount * 2 + 6 + j, 2] = col1[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 3] = col2[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 4] = col3[i * backCount * 2 + j];
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 5], worksheet.Cells[i * backCount * 2 + 7 + j, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[(i * backCount * 2 + j) / 2];
                            Range columnRange8 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 8], worksheet.Cells[i * backCount * 2 + 7 + j, 8]);
                            columnRange8.MergeCells = true;
                            columnRange8.Value2 = col7[(i * backCount * 2 + j) / 2];
                            continue;
                        }
                        worksheet.Cells[i * backCount * 2 + 6 + j, 2] = col1[i * backCount * 2 + j];
                        worksheet.Cells[i * backCount * 2 + 6 + j, 3] = col2[i * backCount * 2 + j];
                        worksheet.Cells[i * backCount * 2 + 6 + j, 4] = col3[i * backCount * 2 + j];
                    }
                }
                worksheet.Cells[col1.Count + 6, 1] = "备注";
                Range rowRemarkRange = worksheet.get_Range(worksheet.Cells[col1.Count + 6, 2], worksheet.Cells[col1.Count + 6, 11]);
                rowRemarkRange.MergeCells = true;
                rowRemarkRange.Value2 = "类型:支导线路线";
            }
            workBook.SaveAs(fileName, fileFormat, nothing, nothing, nothing, nothing, XlSaveAsAccessMode.xlNoChange, nothing, nothing, nothing, nothing, nothing);
            workBook.Close(nothing, nothing, nothing);
            excelApp.Quit();
            foreach (Process p in Process.GetProcessesByName("EXCEL"))
            {
                p.Kill();
            }
        }

        internal void OutputLevelAngle(string path, DataType.Data approximateDataType, int backCount, double angleCloseError, List<string> col0, List<string> col1, List<string> col2, List<string> col3, List<string> col4, List<string> col5, List<string> col6)
        {
            object nothing = Missing.Value;
            object fileFormat = XlFileFormat.xlExcel8;
            object fileName = path;
            Microsoft.Office.Interop.Excel._Application excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Workbook workBook = excelApp.Workbooks.Add(true);
            Worksheet worksheet = workBook.ActiveSheet as Worksheet;
            Range titleRange = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, 7]);
            titleRange.MergeCells = true;
            titleRange.Value2 = "水平角观测记录";
            titleRange.Font.Name = "黑体";
            titleRange.Font.Size = 12;
            titleRange.Font.Bold = true;
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            Range projectNameRange = worksheet.get_Range(worksheet.Cells[2, 1], worksheet.Cells[2, 2]);
            projectNameRange.MergeCells = true;
            projectNameRange.Value2 = "工程名称：";
            Range instrumentRange = worksheet.get_Range(worksheet.Cells[2, 3], worksheet.Cells[2, 4]);
            instrumentRange.MergeCells = true;
            instrumentRange.Value2 = "仪器：";
            Range weatherRange = worksheet.get_Range(worksheet.Cells[2, 5], worksheet.Cells[2, 6]);
            weatherRange.MergeCells = true;
            weatherRange.Value2 = "天气：";
            Range observerRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[3, 2]);
            observerRange.MergeCells = true;
            observerRange.Value2 = "观测者：";
            Range recorderRange = worksheet.get_Range(worksheet.Cells[3, 3], worksheet.Cells[3, 4]);
            recorderRange.MergeCells = true;
            recorderRange.Value2 = "记录者：";
            Range dateRange = worksheet.get_Range(worksheet.Cells[3, 5], worksheet.Cells[3, 6]);
            dateRange.MergeCells = true;
            dateRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");

            Range calculateRange = worksheet.get_Range(worksheet.Cells[col1.Count + 7, 1], worksheet.Cells[col1.Count + 7, 2]);
            calculateRange.MergeCells = true;
            calculateRange.Value2 = "计算者：";
            Range assessmentRange = worksheet.get_Range(worksheet.Cells[col1.Count + 7, 3], worksheet.Cells[col1.Count + 7, 4]);
            assessmentRange.MergeCells = true;
            assessmentRange.Value2 = "审核者：";
            Range endDateRange = worksheet.get_Range(worksheet.Cells[col1.Count + 7, 5], worksheet.Cells[col1.Count + 7, 6]);
            endDateRange.MergeCells = true;
            endDateRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");

            Range borderRange = worksheet.get_Range(worksheet.Cells[4, 1], worksheet.Cells[col1.Count + 6, 7]);
            borderRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, null, null);
            borderRange.Borders.Weight = XlBorderWeight.xlThin;
            Range stationRange = worksheet.get_Range(worksheet.Cells[4, 1], worksheet.Cells[5, 1]);
            stationRange.MergeCells = true;
            stationRange.Value2 = "测站名";
            Range pointRange = worksheet.get_Range(worksheet.Cells[4, 2], worksheet.Cells[5, 2]);
            pointRange.MergeCells = true;
            pointRange.Value2 = "照准点";
            Range levelAngleRange = worksheet.get_Range(worksheet.Cells[4, 3], worksheet.Cells[4, 7]);
            levelAngleRange.MergeCells = true;
            levelAngleRange.Value2 = "水平角(° ' \")";
            worksheet.Cells[5, 3] = "盘左";
            worksheet.Cells[5, 4] = "盘右";
            worksheet.Cells[5, 5] = "2C(\")";
            worksheet.Cells[5, 6] = "角值";
            worksheet.Cells[5, 7] = "方位角";

            if (approximateDataType == DataType.Data.ConnectingTraverse)
            {
                for (int i = 0; i < backCount * 2; i++)
                {
                    if (i == 0)
                    {
                        Range columnRange1 = worksheet.get_Range(worksheet.Cells[i + 6, 1], worksheet.Cells[backCount * 2 + 5, 1]);
                        columnRange1.MergeCells = true;
                        columnRange1.Value2 = col0[i];
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                        columnRange5.MergeCells = true;
                        columnRange5.Value2 = col4[i];
                        Range columnRange6 = worksheet.get_Range(worksheet.Cells[i + 6, 6], worksheet.Cells[backCount * 2 + 5, 6]);
                        columnRange6.MergeCells = true;
                        columnRange6.Value2 = col5[i];
                        Range columnRange7 = worksheet.get_Range(worksheet.Cells[i + 6, 7], worksheet.Cells[backCount + 5, 7]);
                        columnRange7.MergeCells = true;
                        columnRange7.Value2 = col6[i];
                        continue;
                    }
                    if (i == backCount)
                    {
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        if (backCount > 1)
                        {
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[i / 2];
                        }
                        Range columnRange7 = worksheet.get_Range(worksheet.Cells[i + 6, 7], worksheet.Cells[backCount * 2 + 5, 7]);
                        columnRange7.MergeCells = true;
                        columnRange7.Value2 = col6[2];
                        continue;
                    }
                    if (i % 2 == 0 && i != 0 && i != backCount)
                    {
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        if (backCount > 1)
                        {
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[i / 2];
                        }
                        continue;
                    }
                    worksheet.Cells[i + 6, 2] = col1[i];
                    worksheet.Cells[i + 6, 3] = col2[i];
                    worksheet.Cells[i + 6, 4] = col3[i];
                }
                for (int i = 1; i < col1.Count / backCount / 2 - 1; i++)
                {
                    for (int j = 0; j < backCount * 2; j++)
                    {
                        if (j == 0)
                        {
                            Range columnRange1 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 1], worksheet.Cells[(i + 1) * backCount * 2 + 5, 1]);
                            columnRange1.MergeCells = true;
                            columnRange1.Value2 = col0[i];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 2] = col1[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 3] = col2[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 4] = col3[i * backCount * 2 + j];
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 5], worksheet.Cells[i * backCount * 2 + 7 + j, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[(i * backCount * 2 + j) / 2];
                            Range columnRange6 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 6], worksheet.Cells[(i + 1) * backCount * 2 + 5, 6]);
                            columnRange6.MergeCells = true;
                            columnRange6.Value2 = col5[i];
                            Range columnRange7 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 7], worksheet.Cells[(i + 1) * backCount * 2 + 5, 7]);
                            columnRange7.MergeCells = true;
                            columnRange7.Value2 = col6[i + 2];
                            continue;
                        }
                        if (j % 2 == 0 && j != 0)
                        {
                            worksheet.Cells[i * backCount * 2 + 6 + j, 2] = col1[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 3] = col2[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 4] = col3[i * backCount * 2 + j];
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 5], worksheet.Cells[i * backCount * 2 + 7 + j, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[(i * backCount * 2 + j) / 2];
                            continue;
                        }
                        worksheet.Cells[i * backCount * 2 + 6 + j, 2] = col1[i * backCount * 2 + j];
                        worksheet.Cells[i * backCount * 2 + 6 + j, 3] = col2[i * backCount * 2 + j];
                        worksheet.Cells[i * backCount * 2 + 6 + j, 4] = col3[i * backCount * 2 + j];
                    }
                }
                for (int i = col1.Count - backCount * 2; i < col1.Count; i++)
                {
                    if (i == col1.Count - backCount * 2)
                    {
                        Range columnRange1 = worksheet.get_Range(worksheet.Cells[i + 6, 1], worksheet.Cells[i + backCount * 2 + 5, 1]);
                        columnRange1.MergeCells = true;
                        columnRange1.Value2 = col0[i / backCount / 2];
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                        columnRange5.MergeCells = true;
                        columnRange5.Value2 = col4[i / 2];
                        Range columnRange6 = worksheet.get_Range(worksheet.Cells[i + 6, 6], worksheet.Cells[i + backCount * 2 + 5, 6]);
                        columnRange6.MergeCells = true;
                        columnRange6.Value2 = col5[i / backCount / 2];
                        Range columnRange7 = worksheet.get_Range(worksheet.Cells[i + 6, 7], worksheet.Cells[i + backCount * 2 + 5, 7]);
                        columnRange7.MergeCells = true;
                        columnRange7.Value2 = col6[1];
                        continue;
                    }
                    if (i % 2 == 0 && i != col1.Count - backCount * 2)
                    {
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                        columnRange5.MergeCells = true;
                        columnRange5.Value2 = col4[i / 2];
                        continue;
                    }
                    worksheet.Cells[i + 6, 2] = col1[i];
                    worksheet.Cells[i + 6, 3] = col2[i];
                    worksheet.Cells[i + 6, 4] = col3[i];
                }
                worksheet.Cells[col1.Count + 6, 1] = "备注";
                Range rowRemarkRange = worksheet.get_Range(worksheet.Cells[col1.Count + 6, 2], worksheet.Cells[col1.Count + 6, 7]);
                rowRemarkRange.MergeCells = true;
                rowRemarkRange.Value2 = "类型:闭符合路线   角度闭合差:f = " + angleCloseError;
            }
            else if (approximateDataType == DataType.Data.OpenTraverse)
            {
                for (int i = 0; i < backCount * 2; i++)
                {
                    if (i == 0)
                    {
                        Range columnRange1 = worksheet.get_Range(worksheet.Cells[i + 6, 1], worksheet.Cells[backCount * 2 + 5, 1]);
                        columnRange1.MergeCells = true;
                        columnRange1.Value2 = col0[i];
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                        columnRange5.MergeCells = true;
                        columnRange5.Value2 = col4[i];
                        Range columnRange6 = worksheet.get_Range(worksheet.Cells[i + 6, 6], worksheet.Cells[backCount * 2 + 5, 6]);
                        columnRange6.MergeCells = true;
                        columnRange6.Value2 = col5[i];
                        Range columnRange7 = worksheet.get_Range(worksheet.Cells[i + 6, 7], worksheet.Cells[backCount + 5, 7]);
                        columnRange7.MergeCells = true;
                        columnRange7.Value2 = col6[i];
                        continue;
                    }
                    if (i == backCount)
                    {
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        if (backCount > 1)
                        {
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[i / 2];
                        }
                        Range columnRange7 = worksheet.get_Range(worksheet.Cells[i + 6, 7], worksheet.Cells[backCount + 5, 7]);
                        columnRange7.MergeCells = true;
                        columnRange7.Value2 = col6[1];
                        continue;
                    }
                    if (i % 2 == 0 && i != 0 && i != backCount)
                    {
                        worksheet.Cells[i + 6, 2] = col1[i];
                        worksheet.Cells[i + 6, 3] = col2[i];
                        worksheet.Cells[i + 6, 4] = col3[i];
                        if (backCount > 1)
                        {
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i + 6, 5], worksheet.Cells[i + 7, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[i / 2];
                        }
                        continue;
                    }
                    worksheet.Cells[i + 6, 2] = col1[i];
                    worksheet.Cells[i + 6, 3] = col2[i];
                    worksheet.Cells[i + 6, 4] = col3[i];
                }
                for (int i = 1; i < col1.Count / backCount / 2; i++)
                {
                    for (int j = 0; j < backCount * 2; j++)
                    {
                        if (j == 0)
                        {
                            Range columnRange1 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 1], worksheet.Cells[(i + 1) * backCount * 2 + 5, 1]);
                            columnRange1.MergeCells = true;
                            columnRange1.Value2 = col0[i];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 2] = col1[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 3] = col2[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 4] = col3[i * backCount * 2 + j];
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 5], worksheet.Cells[i * backCount * 2 + 7 + j, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[(i * backCount * 2 + j) / 2];
                            Range columnRange6 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 6], worksheet.Cells[(i + 1) * backCount * 2 + 5, 6]);
                            columnRange6.MergeCells = true;
                            columnRange6.Value2 = col5[i];
                            Range columnRange7 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 7], worksheet.Cells[(i + 1) * backCount * 2 + 5, 7]);
                            columnRange7.MergeCells = true;
                            columnRange7.Value2 = col6[i + 1];
                            continue;
                        }
                        if (j % 2 == 0 && j != 0)
                        {
                            worksheet.Cells[i * backCount * 2 + 6 + j, 2] = col1[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 3] = col2[i * backCount * 2 + j];
                            worksheet.Cells[i * backCount * 2 + 6 + j, 4] = col3[i * backCount * 2 + j];
                            Range columnRange5 = worksheet.get_Range(worksheet.Cells[i * backCount * 2 + 6 + j, 5], worksheet.Cells[i * backCount * 2 + 7 + j, 5]);
                            columnRange5.MergeCells = true;
                            columnRange5.Value2 = col4[(i * backCount * 2 + j) / 2];
                            continue;
                        }
                        worksheet.Cells[i * backCount * 2 + 6 + j, 2] = col1[i * backCount * 2 + j];
                        worksheet.Cells[i * backCount * 2 + 6 + j, 3] = col2[i * backCount * 2 + j];
                        worksheet.Cells[i * backCount * 2 + 6 + j, 4] = col3[i * backCount * 2 + j];
                    }
                }
                worksheet.Cells[col1.Count + 6, 1] = "备注";
                Range rowRemarkRange = worksheet.get_Range(worksheet.Cells[col1.Count + 6, 2], worksheet.Cells[col1.Count + 6, 7]);
                rowRemarkRange.MergeCells = true;
                rowRemarkRange.Value2 = "类型:支导线路线";
            }
            workBook.SaveAs(fileName, fileFormat, nothing, nothing, nothing, nothing, XlSaveAsAccessMode.xlNoChange, nothing, nothing, nothing, nothing, nothing);
            workBook.Close(nothing, nothing, nothing);
            excelApp.Quit();
            foreach (Process p in Process.GetProcessesByName("EXCEL"))
            {
                p.Kill();
            }
        }

        internal void OutputPlane(string p, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5, List<string> col8, List<string> col14, List<string> col15)
        {
            object nothing = Missing.Value;
            object fileFormat = XlFileFormat.xlExcel8;
            object fileName = p;
            Microsoft.Office.Interop.Excel._Application excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Workbook workBook = excelApp.Workbooks.Add(true);
            Worksheet worksheet = workBook.ActiveSheet as Worksheet;
            Range titleRange = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, 8]);
            titleRange.MergeCells = true;
            titleRange.Value2 = "导线平差计算";
            titleRange.Font.Name = "黑体";
            titleRange.Font.Size = 12;
            titleRange.Font.Bold = true;
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            Range projectNameRange = worksheet.get_Range(worksheet.Cells[2, 1], worksheet.Cells[2, 2]);
            projectNameRange.MergeCells = true;
            projectNameRange.Value2 = "工程名称：";
            Range instrumentRange = worksheet.get_Range(worksheet.Cells[2, 4], worksheet.Cells[2, 5]);
            instrumentRange.MergeCells = true;
            instrumentRange.Value2 = "仪器：";
            Range weatherRange = worksheet.get_Range(worksheet.Cells[2, 7], worksheet.Cells[2, 8]);
            weatherRange.MergeCells = true;
            weatherRange.Value2 = "天气：";
            Range observerRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[3, 2]);
            observerRange.MergeCells = true;
            observerRange.Value2 = "观测者：";
            Range recorderRange = worksheet.get_Range(worksheet.Cells[3, 4], worksheet.Cells[3, 5]);
            recorderRange.MergeCells = true;
            recorderRange.Value2 = "记录者：";
            Range dateRange = worksheet.get_Range(worksheet.Cells[3, 7], worksheet.Cells[3, 8]);
            dateRange.MergeCells = true;
            dateRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            Range calculateRange = worksheet.get_Range(worksheet.Cells[col0.Count + 6, 1], worksheet.Cells[col0.Count + 6, 2]);
            calculateRange.MergeCells = true;
            calculateRange.Value2 = "计算者：";
            Range assessmentRange = worksheet.get_Range(worksheet.Cells[col0.Count + 6, 4], worksheet.Cells[col0.Count + 6, 5]);
            assessmentRange.MergeCells = true;
            assessmentRange.Value2 = "审核者：";
            Range endDateRange = worksheet.get_Range(worksheet.Cells[col0.Count + 6, 7], worksheet.Cells[col0.Count + 6, 8]);
            endDateRange.MergeCells = true;
            endDateRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            Range borderRange = worksheet.get_Range(worksheet.Cells[4, 1], worksheet.Cells[col1.Count + 5, 8]);
            borderRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, null, null);
            borderRange.Borders.Weight = XlBorderWeight.xlThin;
            Range stationRange = worksheet.get_Range(worksheet.Cells[4, 1], worksheet.Cells[5, 1]);
            stationRange.MergeCells = true;
            stationRange.Value2 = "点名";
            Range samilarCoordRange = worksheet.get_Range(worksheet.Cells[4, 2], worksheet.Cells[4, 3]);
            samilarCoordRange.MergeCells = true;
            samilarCoordRange.Value2 = "近似坐标(m)";
            worksheet.Cells[5, 2] = "X";
            worksheet.Cells[5, 3] = "Y";
            Range samilarSideRange = worksheet.get_Range(worksheet.Cells[4, 4], worksheet.Cells[5, 4]);
            samilarSideRange.MergeCells = true;
            samilarSideRange.Value2 = "近似边长(m)";
            Range samilarAngleRange = worksheet.get_Range(worksheet.Cells[4, 5], worksheet.Cells[5, 5]);
            samilarAngleRange.MergeCells = true;
            samilarAngleRange.Value2 = "近似方位角(° ' \")";
            Range pointErrorRange = worksheet.get_Range(worksheet.Cells[4, 6], worksheet.Cells[5, 6]);
            pointErrorRange.MergeCells = true;
            pointErrorRange.Value2 = "点位中误差(m)";
            Range adjustCoordRange = worksheet.get_Range(worksheet.Cells[4, 7], worksheet.Cells[4, 8]);
            adjustCoordRange.MergeCells = true;
            adjustCoordRange.Value2 = "坐标平差值(m)";
            worksheet.Cells[5, 7] = "X";
            worksheet.Cells[5, 8] = "Y";
            for (int i = 0; i < col0.Count - 1; i++)
            {
                worksheet.Cells[i + 6, 1] = col0[i];
                worksheet.Cells[i + 6, 2] = col1[i];
                worksheet.Cells[i + 6, 3] = col2[i];
                worksheet.Cells[i + 6, 4] = col4[i];
                worksheet.Cells[i + 6, 5] = col5[i];
                worksheet.Cells[i + 6, 6] = col8[i];
                worksheet.Cells[i + 6, 7] = col14[i];
                worksheet.Cells[i + 6, 8] = col15[i];
            }
            worksheet.Cells[col0.Count + 5, 1] = col0[col0.Count - 1];
            worksheet.Cells[col0.Count + 5, 2] = col1[col0.Count - 1];
            worksheet.Cells[col0.Count + 5, 3] = col2[col0.Count - 1];
            worksheet.Cells[col0.Count + 5, 6] = col8[col0.Count - 1];
            worksheet.Cells[col0.Count + 5, 7] = col14[col0.Count - 1];
            worksheet.Cells[col0.Count + 5, 8] = col15[col0.Count - 1];
            workBook.SaveAs(fileName, fileFormat, nothing, nothing, nothing, nothing, XlSaveAsAccessMode.xlNoChange, nothing, nothing, nothing, nothing, nothing);
            workBook.Close(nothing, nothing, nothing);
            excelApp.Quit();
            foreach (Process proc in Process.GetProcessesByName("EXCEL"))
            {
                proc.Kill();
            }
        }

        internal void OutputPoint(string p, List<string> col0, List<string> col14, List<string> col15)
        {
            object nothing = Missing.Value;
            object fileFormat = XlFileFormat.xlExcel8;
            object fileName = p;
            Microsoft.Office.Interop.Excel._Application excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Workbook workBook = excelApp.Workbooks.Add(true);
            Worksheet worksheet = workBook.ActiveSheet as Worksheet;
            Range titleRange = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, 4]);
            titleRange.MergeCells = true;
            titleRange.Value2 = "控制点成果表";
            titleRange.Font.Name = "黑体";
            titleRange.Font.Size = 12;
            titleRange.Font.Bold = true;
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            Range projectNameRange = worksheet.get_Range(worksheet.Cells[2, 1], worksheet.Cells[2, 2]);
            projectNameRange.MergeCells = true;
            projectNameRange.Value2 = "工程名称：";
            Range calculateRange = worksheet.get_Range(worksheet.Cells[2, 3], worksheet.Cells[2, 4]);
            calculateRange.MergeCells = true;
            calculateRange.Value2 = "计算者：";
            Range assessmentRange = worksheet.get_Range(worksheet.Cells[col0.Count + 5, 1], worksheet.Cells[col0.Count + 5, 2]);
            assessmentRange.MergeCells = true;
            assessmentRange.Value2 = "校核者：";
            Range endDateRange = worksheet.get_Range(worksheet.Cells[col0.Count + 5, 3], worksheet.Cells[col0.Count + 5, 4]);
            endDateRange.MergeCells = true;
            endDateRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            Range borderRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[col0.Count + 4, 4]);
            borderRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, null, null);
            borderRange.Borders.Weight = XlBorderWeight.xlThin;
            Range stationRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[4, 1]);
            stationRange.MergeCells = true;
            stationRange.Value2 = "点名";
            Range pointRange = worksheet.get_Range(worksheet.Cells[3, 2], worksheet.Cells[3, 3]);
            pointRange.MergeCells = true;
            pointRange.Value2 = "坐标(m)";
            worksheet.Cells[4, 2] = "X";
            worksheet.Cells[4, 3] = "Y";
            Range demRange = worksheet.get_Range(worksheet.Cells[3, 4], worksheet.Cells[4, 4]);
            demRange.MergeCells = true;
            demRange.Value2 = "高程(m)";
            for (int i = 0; i < col0.Count; i++)
            {
                worksheet.Cells[i + 5, 1] = col0[i];
                worksheet.Cells[i + 5, 2] = col14[i];
                worksheet.Cells[i + 5, 3] = col15[i];
                worksheet.Cells[i + 5, 4] = "0";
            }
            workBook.SaveAs(fileName, fileFormat, nothing, nothing, nothing, nothing, XlSaveAsAccessMode.xlNoChange, nothing, nothing, nothing, nothing, nothing);
            workBook.Close(nothing, nothing, nothing);
            excelApp.Quit();
            foreach (Process proc in Process.GetProcessesByName("EXCEL"))
            {
                proc.Kill();
            }
        }

        internal void OutputResult(string p, List<string> col0, List<string> col14, List<string> col15, List<string> col16, List<string> col17, List<string> col18)
        {
            object nothing = Missing.Value;
            object fileFormat = XlFileFormat.xlExcel8;
            object fileName = p;
            Microsoft.Office.Interop.Excel._Application excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Workbook workBook = excelApp.Workbooks.Add(true);
            Worksheet worksheet = workBook.ActiveSheet as Worksheet;
            Range titleRange = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, 9]);
            titleRange.MergeCells = true;
            titleRange.Value2 = "平差计算成果表";
            titleRange.Font.Name = "黑体";
            titleRange.Font.Size = 12;
            titleRange.Font.Bold = true;
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            Range projectNameRange = worksheet.get_Range(worksheet.Cells[2, 1], worksheet.Cells[2, 4]);
            projectNameRange.MergeCells = true;
            projectNameRange.Value2 = "工程名称：";
            Range calculateRange = worksheet.get_Range(worksheet.Cells[2, 6], worksheet.Cells[2, 9]);
            calculateRange.MergeCells = true;
            calculateRange.Value2 = "计算者：";
            Range assessmentRange = worksheet.get_Range(worksheet.Cells[col0.Count + 5, 1], worksheet.Cells[col0.Count + 5, 4]);
            assessmentRange.MergeCells = true;
            assessmentRange.Value2 = "校核者：";
            Range endDateRange = worksheet.get_Range(worksheet.Cells[col0.Count + 5, 6], worksheet.Cells[col0.Count + 5, 9]);
            endDateRange.MergeCells = true;
            endDateRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            Range borderRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[col0.Count + 4, 9]);
            borderRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, null, null);
            borderRange.Borders.Weight = XlBorderWeight.xlThin;
            Range stationRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[4, 1]);
            stationRange.MergeCells = true;
            stationRange.Value2 = "点名";
            Range pointRange = worksheet.get_Range(worksheet.Cells[3, 2], worksheet.Cells[3, 3]);
            pointRange.MergeCells = true;
            pointRange.Value2 = "坐标(m)";
            worksheet.Cells[4, 2] = "X";
            worksheet.Cells[4, 3] = "Y";
            Range demRange = worksheet.get_Range(worksheet.Cells[3, 4], worksheet.Cells[4, 4]);
            demRange.MergeCells = true;
            demRange.Value2 = "高程(m)";
            Range angleRange = worksheet.get_Range(worksheet.Cells[3, 5], worksheet.Cells[4, 5]);
            angleRange.MergeCells = true;
            angleRange.Value2 = "角度平差值\n(° ' \")";
            Range toPointRange = worksheet.get_Range(worksheet.Cells[3, 6], worksheet.Cells[4, 6]);
            toPointRange.MergeCells = true;
            toPointRange.Value2 = "至点";
            Range dirRange = worksheet.get_Range(worksheet.Cells[3, 7], worksheet.Cells[4, 7]);
            dirRange.MergeCells = true;
            dirRange.Value2 = "方位角\n(° ' \")";
            Range sideRange = worksheet.get_Range(worksheet.Cells[3, 8], worksheet.Cells[4, 8]);
            sideRange.MergeCells = true;
            sideRange.Value2 = "边长平差值\n(m)";
            Range demErrRange = worksheet.get_Range(worksheet.Cells[3, 9], worksheet.Cells[4, 9]);
            demErrRange.MergeCells = true;
            demErrRange.Value2 = "高程平差值\n(m)";
            for (int i = 0; i < col0.Count; i++)
            {
                if (i == 0)
                {
                    worksheet.Cells[i + 5, 1] = col0[i];
                    worksheet.Cells[i + 5, 2] = col14[i];
                    worksheet.Cells[i + 5, 3] = col15[i];
                    worksheet.Cells[i + 5, 4] = "0";
                    worksheet.Cells[i + 5, 5] = " ";
                    worksheet.Cells[i + 5, 6] = col0[i + 1];
                    worksheet.Cells[i + 5, 7] = col16[i];
                    worksheet.Cells[i + 5, 8] = col17[i];
                    worksheet.Cells[i + 5, 9] = "0";
                }
                else if (i == col0.Count - 1)
                {
                    worksheet.Cells[i + 5, 1] = col0[i];
                    worksheet.Cells[i + 5, 2] = col14[i];
                    worksheet.Cells[i + 5, 3] = col15[i];
                    worksheet.Cells[i + 5, 4] = "0";
                    worksheet.Cells[i + 5, 5] = " ";
                    worksheet.Cells[i + 5, 6] = " ";
                    worksheet.Cells[i + 5, 7] = " ";
                    worksheet.Cells[i + 5, 8] = " ";
                    worksheet.Cells[i + 5, 9] = " ";
                }
                else
                {
                    worksheet.Cells[i + 5, 1] = col0[i];
                    worksheet.Cells[i + 5, 2] = col14[i];
                    worksheet.Cells[i + 5, 3] = col15[i];
                    worksheet.Cells[i + 5, 4] = "0";
                    worksheet.Cells[i + 5, 5] = col18[i - 1];
                    worksheet.Cells[i + 5, 6] = col0[i + 1];
                    worksheet.Cells[i + 5, 7] = col16[i];
                    worksheet.Cells[i + 5, 8] = col17[i];
                    worksheet.Cells[i + 5, 9] = "0";
                }
            }
            workBook.SaveAs(fileName, fileFormat, nothing, nothing, nothing, nothing, XlSaveAsAccessMode.xlNoChange, nothing, nothing, nothing, nothing, nothing);
            workBook.Close(nothing, nothing, nothing);
            excelApp.Quit();
            foreach (Process proc in Process.GetProcessesByName("EXCEL"))
            {
                proc.Kill();
            }
        }

        internal void OutputAccuracy(string p, double unitError, List<string> col0, List<string> col6, List<string> col7, List<string> col8, List<string> col9, List<string> col10, List<string> col11, List<string> col12, List<string> col13)
        {
            object nothing = Missing.Value;
            object fileFormat = XlFileFormat.xlExcel8;
            object fileName = p;
            Microsoft.Office.Interop.Excel._Application excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Workbook workBook = excelApp.Workbooks.Add(true);
            Worksheet worksheet = workBook.ActiveSheet as Worksheet;
            Range titleRange = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, 11]);
            titleRange.MergeCells = true;
            titleRange.Value2 = "精度评定表";
            titleRange.Font.Name = "黑体";
            titleRange.Font.Size = 12;
            titleRange.Font.Bold = true;
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            Range projectNameRange = worksheet.get_Range(worksheet.Cells[2, 1], worksheet.Cells[2, 5]);
            projectNameRange.MergeCells = true;
            projectNameRange.Value2 = "工程名称：";
            Range calculateRange = worksheet.get_Range(worksheet.Cells[2, 7], worksheet.Cells[2, 11]);
            calculateRange.MergeCells = true;
            calculateRange.Value2 = "计算者：";
            Range assessmentRange = worksheet.get_Range(worksheet.Cells[col0.Count + 6, 1], worksheet.Cells[col0.Count + 6, 5]);
            assessmentRange.MergeCells = true;
            assessmentRange.Value2 = "校核者：";
            Range endDateRange = worksheet.get_Range(worksheet.Cells[col0.Count + 6, 7], worksheet.Cells[col0.Count + 6, 11]);
            endDateRange.MergeCells = true;
            endDateRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            Range borderRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[col0.Count + 5, 11]);
            borderRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, null, null);
            borderRange.Borders.Weight = XlBorderWeight.xlThin;
            Range stationRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[4, 1]);
            stationRange.MergeCells = true;
            stationRange.Value2 = "点名";
            Range pointRange = worksheet.get_Range(worksheet.Cells[3, 2], worksheet.Cells[3, 4]);
            pointRange.MergeCells = true;
            pointRange.Value2 = "点位中误差(m)";
            worksheet.Cells[4, 2] = "Mx";
            worksheet.Cells[4, 3] = "My";
            worksheet.Cells[4, 4] = "M";
            Range ellipsesRange = worksheet.get_Range(worksheet.Cells[3, 5], worksheet.Cells[3, 7]);
            ellipsesRange.MergeCells = true;
            ellipsesRange.Value2 = "误差椭圆";
            worksheet.Cells[4, 5] = "A(m)";
            worksheet.Cells[4, 6] = "B(m)";
            worksheet.Cells[4, 7] = "F(° ' \")";
            Range demRange = worksheet.get_Range(worksheet.Cells[3, 8], worksheet.Cells[4, 8]);
            demRange.MergeCells = true;
            demRange.Value2 = "高程中误差(m)";
            Range toPointRange = worksheet.get_Range(worksheet.Cells[3, 9], worksheet.Cells[4, 9]);
            toPointRange.MergeCells = true;
            toPointRange.Value2 = "至点";
            Range dirRange = worksheet.get_Range(worksheet.Cells[3, 10], worksheet.Cells[4, 10]);
            dirRange.MergeCells = true;
            dirRange.Value2 = "方位角中误差\n(° ' \")";
            Range sideRange = worksheet.get_Range(worksheet.Cells[3, 11], worksheet.Cells[4, 11]);
            sideRange.MergeCells = true;
            sideRange.Value2 = "边长中误差(m)";
            worksheet.Cells[col0.Count + 5, 1] = "备注";
            Range markRange = worksheet.get_Range(worksheet.Cells[col0.Count + 5, 2], worksheet.Cells[col0.Count + 5, 11]);
            markRange.MergeCells = true;
            markRange.Value2 = "单位权中误差 = " + unitError;
            for (int i = 0; i < col0.Count; i++)
            {
                if (i == 0 || i == 1 || i == col0.Count - 2)
                {
                    worksheet.Cells[i + 5, 1] = col0[i];
                    worksheet.Cells[i + 5, 2] = col6[i];
                    worksheet.Cells[i + 5, 3] = col7[i];
                    worksheet.Cells[i + 5, 4] = col8[i];
                    worksheet.Cells[i + 5, 5] = " ";
                    worksheet.Cells[i + 5, 6] = " ";
                    worksheet.Cells[i + 5, 7] = " ";
                    worksheet.Cells[i + 5, 8] = "0";
                    worksheet.Cells[i + 5, 9] = col0[i + 1];
                    worksheet.Cells[i + 5, 10] = col10[i];
                    worksheet.Cells[i + 5, 11] = col9[i];
                }
                else if (i == col0.Count - 1)
                {
                    worksheet.Cells[i + 5, 1] = col0[i];
                    worksheet.Cells[i + 5, 2] = col6[i];
                    worksheet.Cells[i + 5, 3] = col7[i];
                    worksheet.Cells[i + 5, 4] = col8[i];
                    worksheet.Cells[i + 5, 5] = " ";
                    worksheet.Cells[i + 5, 6] = " ";
                    worksheet.Cells[i + 5, 7] = " ";
                    worksheet.Cells[i + 5, 8] = " ";
                    worksheet.Cells[i + 5, 9] = " ";
                    worksheet.Cells[i + 5, 10] = " ";
                    worksheet.Cells[i + 5, 11] = " ";
                }
                else
                {
                    worksheet.Cells[i + 5, 1] = col0[i];
                    worksheet.Cells[i + 5, 2] = col6[i];
                    worksheet.Cells[i + 5, 3] = col7[i];
                    worksheet.Cells[i + 5, 4] = col8[i];
                    worksheet.Cells[i + 5, 5] = col11[i - 2];
                    worksheet.Cells[i + 5, 6] = col12[i - 2];
                    worksheet.Cells[i + 5, 7] = col13[i - 2];
                    worksheet.Cells[i + 5, 8] = "0";
                    worksheet.Cells[i + 5, 9] = col0[i + 1];
                    worksheet.Cells[i + 5, 10] = col10[i];
                    worksheet.Cells[i + 5, 11] = col9[i];
                }
            }
            workBook.SaveAs(fileName, fileFormat, nothing, nothing, nothing, nothing, XlSaveAsAccessMode.xlNoChange, nothing, nothing, nothing, nothing, nothing);
            workBook.Close(nothing, nothing, nothing);
            excelApp.Quit();
            foreach (Process proc in Process.GetProcessesByName("EXCEL"))
            {
                proc.Kill();
            }
        }

        internal void OutputBLToXY(string p, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5)
        {
            object nothing = Missing.Value;
            object fileFormat = XlFileFormat.xlExcel8;
            object fileName = p;
            Microsoft.Office.Interop.Excel._Application excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Workbook workBook = excelApp.Workbooks.Add(true);
            Worksheet worksheet = workBook.ActiveSheet as Worksheet;
            Range titleRange = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, 5]);
            titleRange.MergeCells = true;
            titleRange.Value2 = "高斯投影正算";
            titleRange.Font.Name = "黑体";
            titleRange.Font.Size = 12;
            titleRange.Font.Bold = true;
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            Range projectNameRange = worksheet.get_Range(worksheet.Cells[2, 1], worksheet.Cells[2, 2]);
            projectNameRange.MergeCells = true;
            projectNameRange.Value2 = "工程名称：";
            Range instrumentRange = worksheet.get_Range(worksheet.Cells[2, 4], worksheet.Cells[2, 5]);
            instrumentRange.MergeCells = true;
            instrumentRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            Range calculateRange = worksheet.get_Range(worksheet.Cells[col0.Count + 5, 1], worksheet.Cells[col0.Count + 5, 2]);
            calculateRange.MergeCells = true;
            calculateRange.Value2 = "计算：";
            worksheet.Cells[col1.Count + 5, 3] = "复核：";
            Range endDateRange = worksheet.get_Range(worksheet.Cells[col0.Count + 5, 4], worksheet.Cells[col0.Count + 5, 5]);
            endDateRange.MergeCells = true;
            endDateRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");

            Range borderRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[col1.Count + 4, 5]);
            borderRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, null, null);
            borderRange.Borders.Weight = XlBorderWeight.xlThin;
            Range stationRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[4, 1]);
            stationRange.MergeCells = true;
            stationRange.Value2 = "点名";
            Range pointRange = worksheet.get_Range(worksheet.Cells[3, 2], worksheet.Cells[3, 3]);
            pointRange.MergeCells = true;
            pointRange.Value2 = "大地坐标(° ' \")";
            worksheet.Cells[4, 2] = "纬度";
            worksheet.Cells[4, 3] = "经度";
            Range levelAngleRange = worksheet.get_Range(worksheet.Cells[3, 4], worksheet.Cells[3, 5]);
            levelAngleRange.MergeCells = true;
            levelAngleRange.Value2 = "高斯坐标(m)";
            worksheet.Cells[4, 4] = "X";
            worksheet.Cells[4, 5] = "Y";
            for (int i = 0; i < col0.Count; i++)
            {
                worksheet.Cells[i + 5, 1] = col0[i];
                worksheet.Cells[i + 5, 2] = col1[i];
                worksheet.Cells[i + 5, 3] = col2[i];
                worksheet.Cells[i + 5, 4] = col4[i];
                worksheet.Cells[i + 5, 5] = col5[i];
            }
            workBook.SaveAs(fileName, fileFormat, nothing, nothing, nothing, nothing, XlSaveAsAccessMode.xlNoChange, nothing, nothing, nothing, nothing, nothing);
            workBook.Close(nothing, nothing, nothing);
            excelApp.Quit();
            foreach (Process proc in Process.GetProcessesByName("EXCEL"))
            {
                proc.Kill();
            }
        }

        internal void OutputXYToBL(string p, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5)
        {
            object nothing = Missing.Value;
            object fileFormat = XlFileFormat.xlExcel8;
            object fileName = p;
            Microsoft.Office.Interop.Excel._Application excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Workbook workBook = excelApp.Workbooks.Add(true);
            Worksheet worksheet = workBook.ActiveSheet as Worksheet;
            Range titleRange = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, 5]);
            titleRange.MergeCells = true;
            titleRange.Value2 = "高斯投影反算";
            titleRange.Font.Name = "黑体";
            titleRange.Font.Size = 12;
            titleRange.Font.Bold = true;
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            Range projectNameRange = worksheet.get_Range(worksheet.Cells[2, 1], worksheet.Cells[2, 2]);
            projectNameRange.MergeCells = true;
            projectNameRange.Value2 = "工程名称：";
            Range instrumentRange = worksheet.get_Range(worksheet.Cells[2, 4], worksheet.Cells[2, 5]);
            instrumentRange.MergeCells = true;
            instrumentRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            Range calculateRange = worksheet.get_Range(worksheet.Cells[col0.Count + 5, 1], worksheet.Cells[col0.Count + 5, 2]);
            calculateRange.MergeCells = true;
            calculateRange.Value2 = "计算：";
            worksheet.Cells[col1.Count + 5, 3] = "复核：";
            Range endDateRange = worksheet.get_Range(worksheet.Cells[col0.Count + 5, 4], worksheet.Cells[col0.Count + 5, 5]);
            endDateRange.MergeCells = true;
            endDateRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");

            Range borderRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[col1.Count + 4, 5]);
            borderRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, null, null);
            borderRange.Borders.Weight = XlBorderWeight.xlThin;
            Range stationRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[4, 1]);
            stationRange.MergeCells = true;
            stationRange.Value2 = "点名";
            Range pointRange = worksheet.get_Range(worksheet.Cells[3, 2], worksheet.Cells[3, 3]);
            pointRange.MergeCells = true;
            pointRange.Value2 = "高斯坐标(m)";
            worksheet.Cells[4, 2] = "X";
            worksheet.Cells[4, 3] = "Y";
            Range levelAngleRange = worksheet.get_Range(worksheet.Cells[3, 4], worksheet.Cells[3, 5]);
            levelAngleRange.MergeCells = true;
            levelAngleRange.Value2 = "大地坐标(° ' \")";
            worksheet.Cells[4, 4] = "纬度";
            worksheet.Cells[4, 5] = "经度";
            for (int i = 0; i < col0.Count; i++)
            {
                worksheet.Cells[i + 5, 1] = col0[i];
                worksheet.Cells[i + 5, 2] = col1[i];
                worksheet.Cells[i + 5, 3] = col2[i];
                worksheet.Cells[i + 5, 4] = col4[i];
                worksheet.Cells[i + 5, 5] = col5[i];
            }
            workBook.SaveAs(fileName, fileFormat, nothing, nothing, nothing, nothing, XlSaveAsAccessMode.xlNoChange, nothing, nothing, nothing, nothing, nothing);
            workBook.Close(nothing, nothing, nothing);
            excelApp.Quit();
            foreach (Process proc in Process.GetProcessesByName("EXCEL"))
            {
                proc.Kill();
            }
        }

        internal void OutputXYToXY(string p, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5)
        {
            object nothing = Missing.Value;
            object fileFormat = XlFileFormat.xlExcel8;
            object fileName = p;
            Microsoft.Office.Interop.Excel._Application excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Workbook workBook = excelApp.Workbooks.Add(true);
            Worksheet worksheet = workBook.ActiveSheet as Worksheet;
            Range titleRange = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, 5]);
            titleRange.MergeCells = true;
            titleRange.Value2 = "坐标换带计算";
            titleRange.Font.Name = "黑体";
            titleRange.Font.Size = 12;
            titleRange.Font.Bold = true;
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            Range projectNameRange = worksheet.get_Range(worksheet.Cells[2, 1], worksheet.Cells[2, 2]);
            projectNameRange.MergeCells = true;
            projectNameRange.Value2 = "工程名称：";
            Range instrumentRange = worksheet.get_Range(worksheet.Cells[2, 4], worksheet.Cells[2, 5]);
            instrumentRange.MergeCells = true;
            instrumentRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            Range calculateRange = worksheet.get_Range(worksheet.Cells[col0.Count + 5, 1], worksheet.Cells[col0.Count + 5, 2]);
            calculateRange.MergeCells = true;
            calculateRange.Value2 = "计算：";
            worksheet.Cells[col1.Count + 5, 3] = "复核：";
            Range endDateRange = worksheet.get_Range(worksheet.Cells[col0.Count + 5, 4], worksheet.Cells[col0.Count + 5, 5]);
            endDateRange.MergeCells = true;
            endDateRange.Value2 = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");

            Range borderRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[col1.Count + 4, 5]);
            borderRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, null, null);
            borderRange.Borders.Weight = XlBorderWeight.xlThin;
            Range stationRange = worksheet.get_Range(worksheet.Cells[3, 1], worksheet.Cells[4, 1]);
            stationRange.MergeCells = true;
            stationRange.Value2 = "点名";
            Range pointRange = worksheet.get_Range(worksheet.Cells[3, 2], worksheet.Cells[3, 3]);
            pointRange.MergeCells = true;
            pointRange.Value2 = "转换前的坐标(m)";
            worksheet.Cells[4, 2] = "X";
            worksheet.Cells[4, 3] = "Y";
            Range levelAngleRange = worksheet.get_Range(worksheet.Cells[3, 4], worksheet.Cells[3, 5]);
            levelAngleRange.MergeCells = true;
            levelAngleRange.Value2 = "转换后的坐标(m)";
            worksheet.Cells[4, 4] = "X";
            worksheet.Cells[4, 5] = "Y";
            for (int i = 0; i < col0.Count; i++)
            {
                worksheet.Cells[i + 5, 1] = col0[i];
                worksheet.Cells[i + 5, 2] = col1[i];
                worksheet.Cells[i + 5, 3] = col2[i];
                worksheet.Cells[i + 5, 4] = col4[i];
                worksheet.Cells[i + 5, 5] = col5[i];
            }
            workBook.SaveAs(fileName, fileFormat, nothing, nothing, nothing, nothing, XlSaveAsAccessMode.xlNoChange, nothing, nothing, nothing, nothing, nothing);
            workBook.Close(nothing, nothing, nothing);
            excelApp.Quit();
            foreach (Process proc in Process.GetProcessesByName("EXCEL"))
            {
                proc.Kill();
            }
        }
    }
}
