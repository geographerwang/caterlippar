using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace AdjustmentAssistant
{
    class OutputWord
    {
        internal void OutputTraverse(string path, DataType.Data approximateDataType, int backCount, double angleCloseError, double coordinateCloseError, double k, List<string> col0, List<string> col1, List<string> col2, List<string> col3, List<string> col4, List<string> col5, List<string> col6, List<string> col7, List<string> col8, List<string> col9, List<string> col10)
        {
            object nothing = Missing.Value;
            object fileFormat = WdSaveFormat.wdFormatDocument;
            object filePath = path;
            Microsoft.Office.Interop.Word._Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
            Microsoft.Office.Interop.Word._Document wordDoc = wordApp.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);
            wordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;
            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter("平差助手");
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
            wordApp.Selection.ParagraphFormat.LineSpacing = 15f;
            object count = col1.Count + 7;
            object wdLine = WdUnits.wdLine;
            for (int i = 0; i < 6; i++)
            {
                wordApp.Selection.MoveDown(ref wdLine, ref count, ref nothing);
                wordApp.Selection.TypeParagraph();
            }
            wordDoc.Paragraphs.First.Range.Text = "导线观测记录";
            wordDoc.Paragraphs.First.Range.Font.Bold = 2;
            wordDoc.Paragraphs.First.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordDoc.Paragraphs[2].Range.Text = string.Format("{0,-27}{1,-29}{2,-28}", "工程名称：", "仪器：", "天气：");
            wordDoc.Paragraphs[2].Range.Font.Bold = 0;
            wordDoc.Paragraphs[2].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordDoc.Paragraphs[3].Range.Text = string.Format("{0,-28}{1,-28}{2,-28}", "观测者：", "记录者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            wordDoc.Paragraphs[3].Range.Font.Bold = 0;
            wordDoc.Paragraphs[3].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordDoc.Paragraphs.Last.Range.Text = string.Format("{0,-28}{1,-28}{2,-28}", "计算者：", "审核者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));

            Table tblTraverse = wordDoc.Tables.Add(wordApp.Selection.Range, col1.Count + 3, 11, ref nothing, ref nothing);
            tblTraverse.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleThickThinLargeGap;
            tblTraverse.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            tblTraverse.Columns[1].Width = 48f;
            tblTraverse.Columns[2].Width = 48f;
            tblTraverse.Columns[3].Width = 60f;
            tblTraverse.Columns[4].Width = 60f;
            tblTraverse.Columns[5].Width = 48f;
            tblTraverse.Columns[6].Width = 60f;
            tblTraverse.Columns[7].Width = 60f;
            tblTraverse.Columns[8].Width = 48f;
            tblTraverse.Columns[9].Width = 48f;
            tblTraverse.Columns[10].Width = 70f;
            tblTraverse.Columns[11].Width = 70f;
            tblTraverse.Cell(1, 1).Range.Text = "测站名";
            tblTraverse.Cell(1, 1).Merge(tblTraverse.Cell(2, 1));
            tblTraverse.Cell(1, 1).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(1, 2).Range.Text = "照准点";
            tblTraverse.Cell(1, 2).Merge(tblTraverse.Cell(2, 2));
            tblTraverse.Cell(1, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(1, 3).Range.Text = "水平角(° '  \")";
            tblTraverse.Cell(1, 3).Merge(tblTraverse.Cell(1, 7));
            tblTraverse.Cell(1, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(2, 3).Range.Text = "盘左读数";
            tblTraverse.Cell(2, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(2, 4).Range.Text = "盘右读数";
            tblTraverse.Cell(2, 4).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(2, 5).Range.Text = "2C(\")";
            tblTraverse.Cell(2, 5).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(2, 6).Range.Text = "角值";
            tblTraverse.Cell(2, 6).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(2, 7).Range.Text = "方位角";
            tblTraverse.Cell(2, 7).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(1, 4).Range.Text = "距离(m)";
            tblTraverse.Cell(1, 4).Merge(tblTraverse.Cell(1, 5));
            tblTraverse.Cell(1, 4).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(2, 8).Range.Text = "平距";
            tblTraverse.Cell(2, 8).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(2, 9).Range.Text = "平均距离";
            tblTraverse.Cell(2, 9).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(1, 5).Range.Text = "坐标(m)";
            tblTraverse.Cell(1, 5).Merge(tblTraverse.Cell(1, 6));
            tblTraverse.Cell(1, 5).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(2, 10).Range.Text = "X";
            tblTraverse.Cell(2, 10).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(2, 11).Range.Text = "Y";
            tblTraverse.Cell(2, 11).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            object moveUnit = WdUnits.wdLine;
            object moveExtend = WdMovementType.wdExtend;
            if (approximateDataType == DataType.Data.ConnectingTraverse)
            {
                for (int i = 0; i < backCount * 2; i++)
                {
                    if (i == 0)
                    {
                        //第一列
                        tblTraverse.Cell(i + 3, 1).Select();
                        tblTraverse.Cell(i + 3, 1).Range.Text = col0[i];
                        object moveCount1 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount1, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        tblTraverse.Cell(i + 3, 5).Select();
                        tblTraverse.Cell(i + 3, 5).Range.Text = col4[i];
                        object moveCount5 = 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第六列
                        tblTraverse.Cell(i + 3, 6).Select();
                        tblTraverse.Cell(i + 3, 6).Range.Text = col5[i];
                        object moveCount6 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount6, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第七列
                        tblTraverse.Cell(i + 3, 7).Select();
                        tblTraverse.Cell(i + 3, 7).Range.Text = col6[i];
                        if (backCount - 1 > 0)
                        {
                            object moveCount7 = backCount - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount7, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        //第八列
                        tblTraverse.Cell(i + 3, 8).Select();
                        tblTraverse.Cell(i + 3, 8).Range.Text = col7[i];
                        object moveCount8 = 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount8, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第九列
                        tblTraverse.Cell(i + 3, 9).Select();
                        tblTraverse.Cell(i + 3, 9).Range.Text = col8[i];
                        object moveCount9 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount9, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第十列
                        tblTraverse.Cell(i + 3, 10).Select();
                        tblTraverse.Cell(i + 3, 10).Range.Text = col9[i];
                        object moveCount10 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount10, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第十一列
                        tblTraverse.Cell(i + 3, 11).Select();
                        tblTraverse.Cell(i + 3, 11).Range.Text = col10[i];
                        object moveCount11 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount11, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        continue;
                    }
                    if (i == backCount)
                    {
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        if (backCount > 1)
                        {
                            tblTraverse.Cell(i + 3, 5).Select();
                            tblTraverse.Cell(i + 3, 5).Range.Text = col4[i / 2];
                            object moveCount5 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        //第七列
                        tblTraverse.Cell(i + 3, 7).Select();
                        tblTraverse.Cell(i + 3, 7).Range.Text = col6[2];
                        if (backCount - 1 > 0)
                        {
                            object moveCount7 = backCount - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount7, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        //第八列
                        if (backCount > 1)
                        {
                            tblTraverse.Cell(i + 3, 8).Select();
                            tblTraverse.Cell(i + 3, 8).Range.Text = col7[i / 2];
                            object moveCount8 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount8, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            continue;
                        }
                    }
                    if (i % 2 == 0 && i != 0 && i != backCount)
                    {
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        tblTraverse.Cell(i + 3, 5).Select();
                        tblTraverse.Cell(i + 3, 5).Range.Text = col4[i / 2];
                        object moveCount5 = 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第八列
                        if (backCount > 1)
                        {
                            tblTraverse.Cell(i + 3, 8).Select();
                            tblTraverse.Cell(i + 3, 8).Range.Text = col7[i / 2];
                            object moveCount8 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount8, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            continue;
                        }
                    }
                    //第二列
                    tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                    //第三列
                    tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                    //第四列
                    tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                }
                for (int i = 1; i < col1.Count / backCount / 2 - 1; i++)
                {
                    for (int j = 0; j < backCount * 2; j++)
                    {
                        if (j == 0)
                        {
                            //第一列
                            tblTraverse.Cell(i * backCount * 2 + 3, 1).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3, 1).Range.Text = col0[i];
                            object moveCount1 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount1, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第二列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 2).Range.Text = col1[i * backCount * 2 + j];
                            //第三列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 3).Range.Text = col2[i * backCount * 2 + j];
                            //第四列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 4).Range.Text = col3[i * backCount * 2 + j];
                            //第五列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Range.Text = col4[(i * backCount * 2 + j) / 2];
                            object moveCount5 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第六列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 6).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 6).Range.Text = col5[i];
                            object moveCount6 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount6, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第七列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 7).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 7).Range.Text = col6[i + 2];
                            if (backCount - 1 > 0)
                            {
                                object moveCount7 = backCount * 2 - 1;
                                wordApp.Selection.MoveDown(ref moveUnit, ref moveCount7, ref moveExtend);
                                wordApp.Selection.Cells.Merge();
                            }
                            //第八列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 8).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 8).Range.Text = col7[(i * backCount * 2 + j) / 2];
                            object moveCount8 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount8, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第九列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 9).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 9).Range.Text = col8[i];
                            object moveCount9 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount9, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第十列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 10).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 10).Range.Text = col9[i + 1];
                            object moveCount10 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount10, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第十一列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 11).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 11).Range.Text = col10[i + 1];
                            object moveCount11 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount11, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            continue;
                        }
                        if (j % 2 == 0 && j != 0)
                        {
                            //第二列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 2).Range.Text = col1[i * backCount * 2 + j];
                            //第三列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 3).Range.Text = col2[i * backCount * 2 + j];
                            //第四列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 4).Range.Text = col3[i * backCount * 2 + j];
                            //第五列
                            if (backCount > 1)
                            {
                                tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Select();
                                tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Range.Text = col4[(i * backCount * 2 + j) / 2];
                                object moveCount5 = 1;
                                wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                                wordApp.Selection.Cells.Merge();
                            }
                            //第八列
                            if (backCount > 1)
                            {
                                tblTraverse.Cell(i * backCount * 2 + 3 + j, 8).Select();
                                tblTraverse.Cell(i * backCount * 2 + 3 + j, 8).Range.Text = col7[(i * backCount * 2 + j) / 2];
                                object moveCount8 = 1;
                                wordApp.Selection.MoveDown(ref moveUnit, ref moveCount8, ref moveExtend);
                                wordApp.Selection.Cells.Merge();
                                continue;
                            }
                        }
                        //第二列
                        tblTraverse.Cell(i * backCount * 2 + 3 + j, 2).Range.Text = col1[i * backCount * 2 + j];
                        //第三列
                        tblTraverse.Cell(i * backCount * 2 + 3 + j, 3).Range.Text = col2[i * backCount * 2 + j];
                        //第四列
                        tblTraverse.Cell(i * backCount * 2 + 3 + j, 4).Range.Text = col3[i * backCount * 2 + j];
                    }
                }
                for (int i = col1.Count - backCount * 2; i < col1.Count; i++)
                {
                    if (i == col1.Count - backCount * 2)
                    {
                        //第一列
                        tblTraverse.Cell(i + 3, 1).Select();
                        tblTraverse.Cell(i + 3, 1).Range.Text = col0[i / backCount / 2];
                        object moveCount1 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount1, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        tblTraverse.Cell(i + 3, 5).Select();
                        tblTraverse.Cell(i + 3, 5).Range.Text = col4[i / 2];
                        object moveCount5 = 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第六列
                        tblTraverse.Cell(i + 3, 6).Select();
                        tblTraverse.Cell(i + 3, 6).Range.Text = col5[i / backCount / 2];
                        object moveCount6 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount6, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第七列
                        tblTraverse.Cell(i + 3, 7).Select();
                        tblTraverse.Cell(i + 3, 7).Range.Text = col6[1];
                        if (backCount - 1 > 0)
                        {
                            object moveCount7 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount7, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        //第八列
                        tblTraverse.Cell(i + 3, 8).Select();
                        tblTraverse.Cell(i + 3, 8).Range.Text = col7[i / 2];
                        object moveCount8 = 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount8, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第九列
                        tblTraverse.Cell(i + 3, 9).Select();
                        tblTraverse.Cell(i + 3, 9).Range.Text = col8[i / backCount / 2];
                        object moveCount9 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount9, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第十列
                        tblTraverse.Cell(i + 3, 10).Select();
                        tblTraverse.Cell(i + 3, 10).Range.Text = col9[1];
                        object moveCount10 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount10, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第十一列
                        tblTraverse.Cell(i + 3, 11).Select();
                        tblTraverse.Cell(i + 3, 11).Range.Text = col10[1];
                        object moveCount11 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount11, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        continue;
                    }
                    if (i % 2 == 0 && i != (col1.Count - backCount * 2))
                    {
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        if (backCount > 1)
                        {
                            tblTraverse.Cell(i + 3, 5).Select();
                            tblTraverse.Cell(i + 3, 5).Range.Text = col4[i / 2];
                            object moveCount5 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        //第八列
                        if (backCount > 1)
                        {
                            tblTraverse.Cell(i + 3, 8).Select();
                            tblTraverse.Cell(i + 3, 8).Range.Text = col7[i / 2];
                            object moveCount8 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount8, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            continue;
                        }
                    }
                    //第二列
                    tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                    //第三列
                    tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                    //第四列
                    tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                }
                tblTraverse.Cell(col1.Count + 3, 1).Range.Text = "备注";
                tblTraverse.Cell(col1.Count + 3, 2).Range.Text = "类型:闭符合路线   角度闭合差:" + angleCloseError + " 坐标增量闭合差:±" + coordinateCloseError + " K≈1/" + k;
                tblTraverse.Cell(col1.Count + 3, 2).Merge(tblTraverse.Cell(col1.Count + 3, 11));
            }
            else if (approximateDataType == DataType.Data.OpenTraverse)
            {
                for (int i = 0; i < backCount * 2; i++)
                {
                    if (i == 0)
                    {
                        //第一列
                        tblTraverse.Cell(i + 3, 1).Select();
                        tblTraverse.Cell(i + 3, 1).Range.Text = col0[i];
                        object moveCount1 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount1, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        tblTraverse.Cell(i + 3, 5).Select();
                        tblTraverse.Cell(i + 3, 5).Range.Text = col4[i];
                        object moveCount5 = 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第六列
                        tblTraverse.Cell(i + 3, 6).Select();
                        tblTraverse.Cell(i + 3, 6).Range.Text = col5[i];
                        object moveCount6 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount6, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第七列
                        tblTraverse.Cell(i + 3, 7).Select();
                        tblTraverse.Cell(i + 3, 7).Range.Text = col6[i];
                        if (backCount - 1 > 0)
                        {
                            object moveCount7 = backCount - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount7, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        //第八列
                        tblTraverse.Cell(i + 3, 8).Select();
                        tblTraverse.Cell(i + 3, 8).Range.Text = col7[i];
                        object moveCount8 = 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount8, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第九列
                        tblTraverse.Cell(i + 3, 9).Select();
                        tblTraverse.Cell(i + 3, 9).Range.Text = col8[i];
                        object moveCount9 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount9, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第十列
                        tblTraverse.Cell(i + 3, 10).Select();
                        tblTraverse.Cell(i + 3, 10).Range.Text = col9[i];
                        object moveCount10 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount10, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第十一列
                        tblTraverse.Cell(i + 3, 11).Select();
                        tblTraverse.Cell(i + 3, 11).Range.Text = col10[i];
                        object moveCount11 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount11, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        continue;
                    }
                    if (i == backCount)
                    {
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        if (backCount > 1)
                        {
                            tblTraverse.Cell(i + 3, 5).Select();
                            tblTraverse.Cell(i + 3, 5).Range.Text = col4[i / 2];
                            object moveCount5 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        //第七列
                        tblTraverse.Cell(i + 3, 7).Select();
                        tblTraverse.Cell(i + 3, 7).Range.Text = col6[1];
                        if (backCount - 1 > 0)
                        {
                            object moveCount7 = backCount - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount7, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        //第八列
                        if (backCount > 1)
                        {
                            tblTraverse.Cell(i + 3, 8).Select();
                            tblTraverse.Cell(i + 3, 8).Range.Text = col7[i / 2];
                            object moveCount8 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount8, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            continue;
                        }
                    }
                    if (i % 2 == 0 && i != 0 && i != backCount)
                    {
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        tblTraverse.Cell(i + 3, 5).Select();
                        tblTraverse.Cell(i + 3, 5).Range.Text = col4[i / 2];
                        object moveCount5 = 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第八列
                        tblTraverse.Cell(i + 3, 8).Select();
                        tblTraverse.Cell(i + 3, 8).Range.Text = col7[i / 2];
                        object moveCount8 = 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount8, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        continue;
                    }
                    //第二列
                    tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                    //第三列
                    tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                    //第四列
                    tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                }
                for (int i = 1; i < col1.Count / backCount / 2; i++)
                {
                    for (int j = 0; j < backCount * 2; j++)
                    {
                        if (j == 0)
                        {
                            //第一列
                            tblTraverse.Cell(i * backCount * 2 + 3, 1).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3, 1).Range.Text = col0[i];
                            object moveCount1 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount1, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第二列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 2).Range.Text = col1[i * backCount * 2 + j];
                            //第三列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 3).Range.Text = col2[i * backCount * 2 + j];
                            //第四列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 4).Range.Text = col3[i * backCount * 2 + j];
                            //第五列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Range.Text = col4[(i * backCount * 2 + j) / 2];
                            object moveCount5 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第六列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 6).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 6).Range.Text = col5[i];
                            object moveCount6 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount6, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第七列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 7).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 7).Range.Text = col6[i + 1];
                            if (backCount - 1 > 0)
                            {
                                object moveCount7 = backCount * 2 - 1;
                                wordApp.Selection.MoveDown(ref moveUnit, ref moveCount7, ref moveExtend);
                                wordApp.Selection.Cells.Merge();
                            }
                            //第八列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 8).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 8).Range.Text = col7[(i * backCount * 2 + j) / 2];
                            object moveCount8 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount8, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第九列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 9).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 9).Range.Text = col8[i];
                            object moveCount9 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount9, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第十列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 10).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 10).Range.Text = col9[i];
                            object moveCount10 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount10, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第十一列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 11).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 11).Range.Text = col10[i];
                            object moveCount11 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount11, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            continue;
                        }
                        if (j % 2 == 0 && j != 0)
                        {
                            //第二列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 2).Range.Text = col1[i * backCount * 2 + j];
                            //第三列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 3).Range.Text = col2[i * backCount * 2 + j];
                            //第四列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 4).Range.Text = col3[i * backCount * 2 + j];
                            //第五列
                            if (backCount > 0)
                            {
                                tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Select();
                                tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Range.Text = col4[(i * backCount * 2 + j) / 2];
                                object moveCount5 = 1;
                                wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                                wordApp.Selection.Cells.Merge();
                            }
                            //第八列
                            if (backCount > 1)
                            {
                                tblTraverse.Cell(i * backCount * 2 + 3 + j, 8).Select();
                                tblTraverse.Cell(i * backCount * 2 + 3 + j, 8).Range.Text = col7[(i * backCount * 2 + j) / 2];
                                object moveCount8 = 1;
                                wordApp.Selection.MoveDown(ref moveUnit, ref moveCount8, ref moveExtend);
                                wordApp.Selection.Cells.Merge();
                                continue;
                            }
                        }
                        //第二列
                        tblTraverse.Cell(i * backCount * 2 + 3 + j, 2).Range.Text = col1[i * backCount * 2 + j];
                        //第三列
                        tblTraverse.Cell(i * backCount * 2 + 3 + j, 3).Range.Text = col2[i * backCount * 2 + j];
                        //第四列
                        tblTraverse.Cell(i * backCount * 2 + 3 + j, 4).Range.Text = col3[i * backCount * 2 + j];
                    }
                }
                tblTraverse.Cell(col1.Count + 3, 1).Range.Text = "备注";
                tblTraverse.Cell(col1.Count + 3, 2).Range.Text = "类型:支导线路线";
                tblTraverse.Cell(col1.Count + 3, 2).Merge(tblTraverse.Cell(col1.Count + 3, 11));
            }
            wordDoc.SaveAs2(ref filePath, ref fileFormat, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing);
            wordDoc.Close(ref nothing, ref nothing, ref nothing);
            wordApp.Quit();
            foreach (Process p in Process.GetProcessesByName("WINWORD"))
            {
                p.Kill();
            }
        }

        internal void OutputLevelAngle(string path, DataType.Data approximateDataType, int backCount, double angleCloseError, List<string> col0, List<string> col1, List<string> col2, List<string> col3, List<string> col4, List<string> col5, List<string> col6)
        {
            object nothing = Missing.Value;
            object fileFormat = WdSaveFormat.wdFormatDocument;
            object filePath = path;
            Microsoft.Office.Interop.Word._Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
            Microsoft.Office.Interop.Word._Document wordDoc = wordApp.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);
            wordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;
            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter("平差助手");
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
            wordApp.Selection.ParagraphFormat.LineSpacing = 15f;
            object count = col1.Count + 7;
            object wdLine = WdUnits.wdLine;
            for (int i = 0; i < 6; i++)
            {
                wordApp.Selection.MoveDown(ref wdLine, ref count, ref nothing);
                wordApp.Selection.TypeParagraph();
            }
            wordDoc.Paragraphs.First.Range.Text = "";
            wordDoc.Paragraphs.First.Range.Font.Bold = 2;
            wordDoc.Paragraphs.First.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordDoc.Paragraphs[2].Range.Text = string.Format("{0,-27}{1,-29}{2,-28}", "工程名称：", "仪器：", "天气：");
            wordDoc.Paragraphs[2].Range.Font.Bold = 0;
            wordDoc.Paragraphs[2].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordDoc.Paragraphs[3].Range.Text = string.Format("{0,-28}{1,-28}{2,-28}", "观测者：", "记录者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            wordDoc.Paragraphs[3].Range.Font.Bold = 0;
            wordDoc.Paragraphs[3].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordDoc.Paragraphs.Last.Range.Text = string.Format("{0,-28}{1,-28}{2,-28}", "计算者：", "审核者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));

            Table tblTraverse = wordDoc.Tables.Add(wordApp.Selection.Range, col1.Count + 3, 7, ref nothing, ref nothing);
            tblTraverse.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleThickThinLargeGap;
            tblTraverse.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            tblTraverse.Columns[1].Width = 48f;
            tblTraverse.Columns[2].Width = 48f;
            tblTraverse.Columns[3].Width = 60f;
            tblTraverse.Columns[4].Width = 60f;
            tblTraverse.Columns[5].Width = 48f;
            tblTraverse.Columns[6].Width = 60f;
            tblTraverse.Columns[7].Width = 60f;
            tblTraverse.Cell(1, 1).Range.Text = "测站名";
            tblTraverse.Cell(1, 1).Merge(tblTraverse.Cell(2, 1));
            tblTraverse.Cell(1, 1).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(1, 2).Range.Text = "照准点";
            tblTraverse.Cell(1, 2).Merge(tblTraverse.Cell(2, 2));
            tblTraverse.Cell(1, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(1, 3).Range.Text = "水平角(° '  \")";
            tblTraverse.Cell(1, 3).Merge(tblTraverse.Cell(1, 7));
            tblTraverse.Cell(1, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(2, 3).Range.Text = "盘左读数";
            tblTraverse.Cell(2, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(2, 4).Range.Text = "盘右读数";
            tblTraverse.Cell(2, 4).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(2, 5).Range.Text = "2C(\")";
            tblTraverse.Cell(2, 5).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(2, 6).Range.Text = "角值";
            tblTraverse.Cell(2, 6).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblTraverse.Cell(2, 7).Range.Text = "方位角";
            tblTraverse.Cell(2, 7).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            object moveUnit = WdUnits.wdLine;
            object moveExtend = WdMovementType.wdExtend;
            if (approximateDataType == DataType.Data.ConnectingTraverse)
            {
                for (int i = 0; i < backCount * 2; i++)
                {
                    if (i == 0)
                    {
                        //第一列
                        tblTraverse.Cell(i + 3, 1).Select();
                        tblTraverse.Cell(i + 3, 1).Range.Text = col0[i];
                        object moveCount1 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount1, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        tblTraverse.Cell(i + 3, 5).Select();
                        tblTraverse.Cell(i + 3, 5).Range.Text = col4[i];
                        object moveCount5 = 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第六列
                        tblTraverse.Cell(i + 3, 6).Select();
                        tblTraverse.Cell(i + 3, 6).Range.Text = col5[i];
                        object moveCount6 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount6, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第七列
                        tblTraverse.Cell(i + 3, 7).Select();
                        tblTraverse.Cell(i + 3, 7).Range.Text = col6[i];
                        if (backCount - 1 > 0)
                        {
                            object moveCount7 = backCount - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount7, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        continue;
                    }
                    if (i == backCount)
                    {
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        if (backCount > 1)
                        {
                            tblTraverse.Cell(i + 3, 5).Select();
                            tblTraverse.Cell(i + 3, 5).Range.Text = col4[i / 2];
                            object moveCount5 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        //第七列
                        tblTraverse.Cell(i + 3, 7).Select();
                        tblTraverse.Cell(i + 3, 7).Range.Text = col6[2];
                        if (backCount - 1 > 0)
                        {
                            object moveCount7 = backCount - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount7, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        continue;
                    }
                    if (i % 2 == 0 && i != 0 && i != backCount)
                    {
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        tblTraverse.Cell(i + 3, 5).Select();
                        tblTraverse.Cell(i + 3, 5).Range.Text = col4[i / 2];
                        object moveCount5 = 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        continue;
                    }
                    //第二列
                    tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                    //第三列
                    tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                    //第四列
                    tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                }
                for (int i = 1; i < col1.Count / backCount / 2 - 1; i++)
                {
                    for (int j = 0; j < backCount * 2; j++)
                    {
                        if (j == 0)
                        {
                            //第一列
                            tblTraverse.Cell(i * backCount * 2 + 3, 1).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3, 1).Range.Text = col0[i];
                            object moveCount1 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount1, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第二列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 2).Range.Text = col1[i * backCount * 2 + j];
                            //第三列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 3).Range.Text = col2[i * backCount * 2 + j];
                            //第四列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 4).Range.Text = col3[i * backCount * 2 + j];
                            //第五列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Range.Text = col4[(i * backCount * 2 + j) / 2];
                            object moveCount5 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第六列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 6).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 6).Range.Text = col5[i];
                            object moveCount6 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount6, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第七列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 7).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 7).Range.Text = col6[i + 2];
                            if (backCount - 1 > 0)
                            {
                                object moveCount7 = backCount * 2 - 1;
                                wordApp.Selection.MoveDown(ref moveUnit, ref moveCount7, ref moveExtend);
                                wordApp.Selection.Cells.Merge();
                            }
                            continue;
                        }
                        if (j % 2 == 0 && j != 0)
                        {
                            //第二列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 2).Range.Text = col1[i * backCount * 2 + j];
                            //第三列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 3).Range.Text = col2[i * backCount * 2 + j];
                            //第四列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 4).Range.Text = col3[i * backCount * 2 + j];
                            //第五列
                            if (backCount > 1)
                            {
                                tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Select();
                                tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Range.Text = col4[(i * backCount * 2 + j) / 2];
                                object moveCount5 = 1;
                                wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                                wordApp.Selection.Cells.Merge();
                            }
                            continue;
                        }
                        //第二列
                        tblTraverse.Cell(i * backCount * 2 + 3 + j, 2).Range.Text = col1[i * backCount * 2 + j];
                        //第三列
                        tblTraverse.Cell(i * backCount * 2 + 3 + j, 3).Range.Text = col2[i * backCount * 2 + j];
                        //第四列
                        tblTraverse.Cell(i * backCount * 2 + 3 + j, 4).Range.Text = col3[i * backCount * 2 + j];
                    }
                }
                for (int i = col1.Count - backCount * 2; i < col1.Count; i++)
                {
                    if (i == col1.Count - backCount * 2)
                    {
                        //第一列
                        tblTraverse.Cell(i + 3, 1).Select();
                        tblTraverse.Cell(i + 3, 1).Range.Text = col0[i / backCount / 2];
                        object moveCount1 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount1, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        tblTraverse.Cell(i + 3, 5).Select();
                        tblTraverse.Cell(i + 3, 5).Range.Text = col4[i / 2];
                        object moveCount5 = 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第六列
                        tblTraverse.Cell(i + 3, 6).Select();
                        tblTraverse.Cell(i + 3, 6).Range.Text = col5[i / backCount / 2];
                        object moveCount6 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount6, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第七列
                        tblTraverse.Cell(i + 3, 7).Select();
                        tblTraverse.Cell(i + 3, 7).Range.Text = col6[1];
                        if (backCount - 1 > 0)
                        {
                            object moveCount7 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount7, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        continue;
                    }
                    if (i % 2 == 0 && i != (col1.Count - backCount * 2))
                    {
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        if (backCount > 1)
                        {
                            tblTraverse.Cell(i + 3, 5).Select();
                            tblTraverse.Cell(i + 3, 5).Range.Text = col4[i / 2];
                            object moveCount5 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        continue;
                    }
                    //第二列
                    tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                    //第三列
                    tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                    //第四列
                    tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                }
                tblTraverse.Cell(col1.Count + 3, 1).Range.Text = "备注";
                tblTraverse.Cell(col1.Count + 3, 2).Range.Text = "类型:闭符合路线   角度闭合差:f = " + angleCloseError;
                tblTraverse.Cell(col1.Count + 3, 2).Merge(tblTraverse.Cell(col1.Count + 3, 7));
            }
            else if (approximateDataType == DataType.Data.OpenTraverse)
            {
                for (int i = 0; i < backCount * 2; i++)
                {
                    if (i == 0)
                    {
                        //第一列
                        tblTraverse.Cell(i + 3, 1).Select();
                        tblTraverse.Cell(i + 3, 1).Range.Text = col0[i];
                        object moveCount1 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount1, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        tblTraverse.Cell(i + 3, 5).Select();
                        tblTraverse.Cell(i + 3, 5).Range.Text = col4[i];
                        object moveCount5 = 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第六列
                        tblTraverse.Cell(i + 3, 6).Select();
                        tblTraverse.Cell(i + 3, 6).Range.Text = col5[i];
                        object moveCount6 = backCount * 2 - 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount6, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        //第七列
                        tblTraverse.Cell(i + 3, 7).Select();
                        tblTraverse.Cell(i + 3, 7).Range.Text = col6[i];
                        if (backCount - 1 > 0)
                        {
                            object moveCount7 = backCount - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount7, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        continue;
                    }
                    if (i == backCount)
                    {
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        if (backCount > 1)
                        {
                            tblTraverse.Cell(i + 3, 5).Select();
                            tblTraverse.Cell(i + 3, 5).Range.Text = col4[i / 2];
                            object moveCount5 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        //第七列
                        tblTraverse.Cell(i + 3, 7).Select();
                        tblTraverse.Cell(i + 3, 7).Range.Text = col6[1];
                        if (backCount - 1 > 0)
                        {
                            object moveCount7 = backCount - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount7, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                        }
                        continue;
                    }
                    if (i % 2 == 0 && i != 0 && i != backCount)
                    {
                        //第二列
                        tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                        //第三列
                        tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                        //第四列
                        tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                        //第五列
                        tblTraverse.Cell(i + 3, 5).Select();
                        tblTraverse.Cell(i + 3, 5).Range.Text = col4[i / 2];
                        object moveCount5 = 1;
                        wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                        wordApp.Selection.Cells.Merge();
                        continue;
                    }
                    //第二列
                    tblTraverse.Cell(i + 3, 2).Range.Text = col1[i];
                    //第三列
                    tblTraverse.Cell(i + 3, 3).Range.Text = col2[i];
                    //第四列
                    tblTraverse.Cell(i + 3, 4).Range.Text = col3[i];
                }
                for (int i = 1; i < col1.Count / backCount / 2; i++)
                {
                    for (int j = 0; j < backCount * 2; j++)
                    {
                        if (j == 0)
                        {
                            //第一列
                            tblTraverse.Cell(i * backCount * 2 + 3, 1).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3, 1).Range.Text = col0[i];
                            object moveCount1 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount1, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第二列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 2).Range.Text = col1[i * backCount * 2 + j];
                            //第三列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 3).Range.Text = col2[i * backCount * 2 + j];
                            //第四列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 4).Range.Text = col3[i * backCount * 2 + j];
                            //第五列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Range.Text = col4[(i * backCount * 2 + j) / 2];
                            object moveCount5 = 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第六列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 6).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 6).Range.Text = col5[i];
                            object moveCount6 = backCount * 2 - 1;
                            wordApp.Selection.MoveDown(ref moveUnit, ref moveCount6, ref moveExtend);
                            wordApp.Selection.Cells.Merge();
                            //第七列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 7).Select();
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 7).Range.Text = col6[i + 1];
                            if (backCount - 1 > 0)
                            {
                                object moveCount7 = backCount * 2 - 1;
                                wordApp.Selection.MoveDown(ref moveUnit, ref moveCount7, ref moveExtend);
                                wordApp.Selection.Cells.Merge();
                            }
                            continue;
                        }
                        if (j % 2 == 0 && j != 0)
                        {
                            //第二列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 2).Range.Text = col1[i * backCount * 2 + j];
                            //第三列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 3).Range.Text = col2[i * backCount * 2 + j];
                            //第四列
                            tblTraverse.Cell(i * backCount * 2 + 3 + j, 4).Range.Text = col3[i * backCount * 2 + j];
                            //第五列
                            if (backCount > 0)
                            {
                                tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Select();
                                tblTraverse.Cell(i * backCount * 2 + 3 + j, 5).Range.Text = col4[(i * backCount * 2 + j) / 2];
                                object moveCount5 = 1;
                                wordApp.Selection.MoveDown(ref moveUnit, ref moveCount5, ref moveExtend);
                                wordApp.Selection.Cells.Merge();
                            }
                            continue;
                        }
                        //第二列
                        tblTraverse.Cell(i * backCount * 2 + 3 + j, 2).Range.Text = col1[i * backCount * 2 + j];
                        //第三列
                        tblTraverse.Cell(i * backCount * 2 + 3 + j, 3).Range.Text = col2[i * backCount * 2 + j];
                        //第四列
                        tblTraverse.Cell(i * backCount * 2 + 3 + j, 4).Range.Text = col3[i * backCount * 2 + j];
                    }
                }
                tblTraverse.Cell(col1.Count + 3, 1).Range.Text = "备注";
                tblTraverse.Cell(col1.Count + 3, 2).Range.Text = "类型:支导线路线";
                tblTraverse.Cell(col1.Count + 3, 2).Merge(tblTraverse.Cell(col1.Count + 3, 7));
            }
            wordDoc.SaveAs2(ref filePath, ref fileFormat, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing);
            wordDoc.Close(ref nothing, ref nothing, ref nothing);
            wordApp.Quit();
            foreach (Process p in Process.GetProcessesByName("WINWORD"))
            {
                p.Kill();
            }
        }

        internal void OutputPlane(string p, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5, List<string> col8, List<string> col14, List<string> col15)
        {
            object nothing = Missing.Value;
            object fileFormat = WdSaveFormat.wdFormatDocument;
            object filePath = p;
            Microsoft.Office.Interop.Word._Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
            Microsoft.Office.Interop.Word._Document wordDoc = wordApp.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);
            wordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;
            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter("平差助手");
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
            wordApp.Selection.ParagraphFormat.LineSpacing = 15f;
            object count = col1.Count + 6;
            object wdLine = WdUnits.wdLine;
            for (int i = 0; i < 6; i++)
            {
                wordApp.Selection.MoveDown(ref wdLine, ref count, ref nothing);
                wordApp.Selection.TypeParagraph();
            }
            wordDoc.Paragraphs.First.Range.Text = "导线平差计算";
            wordDoc.Paragraphs.First.Range.Font.Bold = 2;
            wordDoc.Paragraphs.First.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordDoc.Paragraphs[2].Range.Text = string.Format("{0,-27}{1,-29}{2,-28}", "工程名称：", "仪器：", "天气：");
            wordDoc.Paragraphs[2].Range.Font.Bold = 0;
            wordDoc.Paragraphs[2].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordDoc.Paragraphs[3].Range.Text = string.Format("{0,-28}{1,-28}{2,-28}", "观测者：", "记录者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            wordDoc.Paragraphs[3].Range.Font.Bold = 0;
            wordDoc.Paragraphs[3].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordDoc.Paragraphs.Last.Range.Text = string.Format("{0,-28}{1,-28}{2,-28}", "计算者：", "审核者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            Table tblPlane = wordDoc.Tables.Add(wordApp.Selection.Range, col1.Count + 2, 8, ref nothing, ref nothing);
            tblPlane.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleThickThinLargeGap;
            tblPlane.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            tblPlane.Columns[1].Width = 30f;
            tblPlane.Columns[2].Width = 80f;
            tblPlane.Columns[3].Width = 80f;
            tblPlane.Columns[4].Width = 60f;
            tblPlane.Columns[5].Width = 60f;
            tblPlane.Columns[6].Width = 70f;
            tblPlane.Columns[7].Width = 95f;
            tblPlane.Columns[8].Width = 95f;
            tblPlane.Cell(1, 1).Range.Text = "点名";
            tblPlane.Cell(1, 1).Merge(tblPlane.Cell(2, 1));
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblPlane.Cell(1, 2).Range.Text = "近似坐标(m)";
            tblPlane.Cell(1, 2).Merge(tblPlane.Cell(1, 3));
            tblPlane.Cell(1, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblPlane.Cell(2, 2).Range.Text = "X";
            tblPlane.Cell(2, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblPlane.Cell(2, 3).Range.Text = "Y";
            tblPlane.Cell(2, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblPlane.Cell(1, 3).Range.Text = "近似边长\n(m)";
            tblPlane.Cell(1, 3).Merge(tblPlane.Cell(2, 4));
            tblPlane.Cell(1, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblPlane.Cell(1, 4).Range.Text = "近似方位角(° ' \")";
            tblPlane.Cell(1, 4).Merge(tblPlane.Cell(2, 5));
            tblPlane.Cell(1, 4).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblPlane.Cell(1, 5).Range.Text = "点位中误差\n(m)";
            tblPlane.Cell(1, 5).Merge(tblPlane.Cell(2, 6));
            tblPlane.Cell(1, 5).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblPlane.Cell(1, 6).Range.Text = "坐标平差值(m)";
            tblPlane.Cell(1, 6).Merge(tblPlane.Cell(1, 7));
            tblPlane.Cell(1, 6).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblPlane.Cell(2, 7).Range.Text = "X";
            tblPlane.Cell(2, 7).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblPlane.Cell(2, 8).Range.Text = "Y";
            tblPlane.Cell(2, 8).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            object moveUnit = WdUnits.wdLine;
            object moveExtend = WdMovementType.wdExtend;
            for (int i = 0; i < col0.Count - 1; i++)
            {
                tblPlane.Cell(i + 3, 1).Range.Text = col0[i];
                tblPlane.Cell(i + 3, 2).Range.Text = col1[i];
                tblPlane.Cell(i + 3, 3).Range.Text = col2[i];
                tblPlane.Cell(i + 3, 4).Range.Text = col4[i];
                tblPlane.Cell(i + 3, 5).Range.Text = col5[i];
                tblPlane.Cell(i + 3, 6).Range.Text = col8[i];
                tblPlane.Cell(i + 3, 7).Range.Text = col14[i];
                tblPlane.Cell(i + 3, 8).Range.Text = col15[i];
            }
            tblPlane.Cell(col0.Count + 2, 1).Range.Text = col0[col0.Count - 1];
            tblPlane.Cell(col0.Count + 2, 2).Range.Text = col1[col0.Count - 1];
            tblPlane.Cell(col0.Count + 2, 3).Range.Text = col2[col0.Count - 1];
            tblPlane.Cell(col0.Count + 2, 6).Range.Text = col8[col0.Count - 1];
            tblPlane.Cell(col0.Count + 2, 7).Range.Text = col14[col0.Count - 1];
            tblPlane.Cell(col0.Count + 2, 8).Range.Text = col15[col0.Count - 1];
            wordDoc.SaveAs2(ref filePath, ref fileFormat, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing);
            wordDoc.Close(ref nothing, ref nothing, ref nothing);
            wordApp.Quit();
            foreach (Process proc in Process.GetProcessesByName("WINWORD"))
            {
                proc.Kill();
            }
        }

        internal void OutputPoint(string p, List<string> col0, List<string> col14, List<string> col15)
        {
            object nothing = Missing.Value;
            object fileFormat = WdSaveFormat.wdFormatDocument;
            object filePath = p;
            Microsoft.Office.Interop.Word._Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
            Microsoft.Office.Interop.Word._Document wordDoc = wordApp.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);
            wordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;
            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter("平差助手");
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
            wordApp.Selection.ParagraphFormat.LineSpacing = 15f;
            object count = col0.Count + 5;
            object wdLine = WdUnits.wdLine;
            for (int i = 0; i < 4; i++)
            {
                wordApp.Selection.MoveDown(ref wdLine, ref count, ref nothing);
                wordApp.Selection.TypeParagraph();
            }
            wordDoc.Paragraphs.First.Range.Text = "控制点成果表";
            wordDoc.Paragraphs.First.Range.Font.Bold = 2;
            wordDoc.Paragraphs.First.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordDoc.Paragraphs[2].Range.Text = string.Format("{0,-25}{1,-25}", "工程名称：", "计算者：");
            wordDoc.Paragraphs[2].Range.Font.Bold = 0;
            wordDoc.Paragraphs[2].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordDoc.Paragraphs.Last.Range.Text = string.Format("{0,-25}{1,-25}", "校核者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));

            Table tblPoint = wordDoc.Tables.Add(wordApp.Selection.Range, col0.Count + 2, 4, ref nothing, ref nothing);
            tblPoint.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleThickThinLargeGap;
            tblPoint.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            tblPoint.Columns[1].Width = 60f;
            tblPoint.Columns[2].Width = 100f;
            tblPoint.Columns[3].Width = 100f;
            tblPoint.Columns[4].Width = 80f;
            tblPoint.Cell(1, 1).Range.Text = "点名";
            tblPoint.Cell(1, 1).Merge(tblPoint.Cell(2, 1));
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblPoint.Cell(1, 2).Range.Text = "坐标(m)";
            tblPoint.Cell(1, 2).Merge(tblPoint.Cell(1, 3));
            tblPoint.Cell(1, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblPoint.Cell(2, 2).Range.Text = "X";
            tblPoint.Cell(2, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblPoint.Cell(2, 3).Range.Text = "Y";
            tblPoint.Cell(2, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblPoint.Cell(1, 3).Range.Text = "高程(m)";
            tblPoint.Cell(1, 3).Merge(tblPoint.Cell(2, 4));
            tblPoint.Cell(1, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            object moveUnit = WdUnits.wdLine;
            object moveExtend = WdMovementType.wdExtend;
            for (int i = 0; i < col0.Count; i++)
            {
                tblPoint.Cell(i + 3, 1).Range.Text = col0[i];
                tblPoint.Cell(i + 3, 2).Range.Text = col14[i];
                tblPoint.Cell(i + 3, 3).Range.Text = col15[i];
                tblPoint.Cell(i + 3, 4).Range.Text = "0";
            }

            wordDoc.SaveAs2(ref filePath, ref fileFormat, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing);
            wordDoc.Close(ref nothing, ref nothing, ref nothing);
            wordApp.Quit();
            foreach (Process proc in Process.GetProcessesByName("WINWORD"))
            {
                proc.Kill();
            }
        }

        internal void OutputResult(string p, List<string> col0, List<string> col14, List<string> col15, List<string> col16, List<string> col17, List<string> col18)
        {
            object nothing = Missing.Value;
            object fileFormat = WdSaveFormat.wdFormatDocument;
            object filePath = p;
            Microsoft.Office.Interop.Word._Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
            Microsoft.Office.Interop.Word._Document wordDoc = wordApp.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);
            wordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;
            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter("平差助手");
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
            wordApp.Selection.ParagraphFormat.LineSpacing = 15f;
            object count = col0.Count + 5;
            object wdLine = WdUnits.wdLine;
            for (int i = 0; i < 4; i++)
            {
                wordApp.Selection.MoveDown(ref wdLine, ref count, ref nothing);
                wordApp.Selection.TypeParagraph();
            }
            wordDoc.Paragraphs.First.Range.Text = "平差计算成果表";
            wordDoc.Paragraphs.First.Range.Font.Bold = 2;
            wordDoc.Paragraphs.First.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordDoc.Paragraphs[2].Range.Text = string.Format("{0,-50}{1,-50}", "工程名称：", "计算者：");
            wordDoc.Paragraphs[2].Range.Font.Bold = 0;
            wordDoc.Paragraphs[2].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordDoc.Paragraphs.Last.Range.Text = string.Format("{0,-50}{1,-50}", "校核者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            Table tblResult = wordDoc.Tables.Add(wordApp.Selection.Range, col0.Count + 2, 9, ref nothing, ref nothing);
            tblResult.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleThickThinLargeGap;
            tblResult.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            tblResult.Columns[1].Width = 30f;
            tblResult.Columns[2].Width = 80f;
            tblResult.Columns[3].Width = 80f;
            tblResult.Columns[4].Width = 50f;
            tblResult.Columns[5].Width = 100f;
            tblResult.Columns[6].Width = 30f;
            tblResult.Columns[7].Width = 100f;
            tblResult.Columns[8].Width = 60f;
            tblResult.Columns[9].Width = 50f;
            tblResult.Cell(1, 1).Range.Text = "点名";
            tblResult.Cell(1, 1).Merge(tblResult.Cell(2, 1));
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblResult.Cell(1, 2).Range.Text = "坐标(m)";
            tblResult.Cell(1, 2).Merge(tblResult.Cell(1, 3));
            tblResult.Cell(1, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblResult.Cell(2, 2).Range.Text = "X";
            tblResult.Cell(2, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblResult.Cell(2, 3).Range.Text = "Y";
            tblResult.Cell(2, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblResult.Cell(1, 3).Range.Text = "高程(m)";
            tblResult.Cell(1, 3).Merge(tblResult.Cell(2, 4));
            tblResult.Cell(1, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblResult.Cell(1, 4).Range.Text = "角度平差值\n(° ' \")";
            tblResult.Cell(1, 4).Merge(tblResult.Cell(2, 5));
            tblResult.Cell(1, 4).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblResult.Cell(1, 5).Range.Text = "至点";
            tblResult.Cell(1, 5).Merge(tblResult.Cell(2, 6));
            tblResult.Cell(1, 5).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblResult.Cell(1, 6).Range.Text = "方位角\n(° ' \")";
            tblResult.Cell(1, 6).Merge(tblResult.Cell(2, 7));
            tblResult.Cell(1, 6).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblResult.Cell(1, 7).Range.Text = "边长平差值(m)";
            tblResult.Cell(1, 7).Merge(tblResult.Cell(2, 8));
            tblResult.Cell(1, 7).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblResult.Cell(1, 8).Range.Text = "高差平差值(m)";
            tblResult.Cell(1, 8).Merge(tblResult.Cell(2, 9));
            tblResult.Cell(1, 8).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            object moveUnit = WdUnits.wdLine;
            object moveExtend = WdMovementType.wdExtend;
            for (int i = 0; i < col0.Count; i++)
            {
                if (i == 0)
                {
                    tblResult.Cell(i + 3, 1).Range.Text = col0[i];
                    tblResult.Cell(i + 3, 2).Range.Text = col14[i];
                    tblResult.Cell(i + 3, 3).Range.Text = col15[i];
                    tblResult.Cell(i + 3, 4).Range.Text = "0";
                    tblResult.Cell(i + 3, 5).Range.Text = " ";
                    tblResult.Cell(i + 3, 6).Range.Text = col0[i + 1];
                    tblResult.Cell(i + 3, 7).Range.Text = col16[i];
                    tblResult.Cell(i + 3, 8).Range.Text = col17[i];
                    tblResult.Cell(i + 3, 9).Range.Text = "0";
                }
                else if (i == col0.Count - 1)
                {
                    tblResult.Cell(i + 3, 1).Range.Text = col0[i];
                    tblResult.Cell(i + 3, 2).Range.Text = col14[i];
                    tblResult.Cell(i + 3, 3).Range.Text = col15[i];
                    tblResult.Cell(i + 3, 4).Range.Text = "0";
                    tblResult.Cell(i + 3, 5).Range.Text = " ";
                    tblResult.Cell(i + 3, 6).Range.Text = " ";
                    tblResult.Cell(i + 3, 7).Range.Text = " ";
                    tblResult.Cell(i + 3, 8).Range.Text = " ";
                    tblResult.Cell(i + 3, 9).Range.Text = " ";
                }
                else
                {
                    tblResult.Cell(i + 3, 1).Range.Text = col0[i];
                    tblResult.Cell(i + 3, 2).Range.Text = col14[i];
                    tblResult.Cell(i + 3, 3).Range.Text = col15[i];
                    tblResult.Cell(i + 3, 4).Range.Text = "0";
                    tblResult.Cell(i + 3, 5).Range.Text = col18[i - 1];
                    tblResult.Cell(i + 3, 6).Range.Text = col0[i + 1];
                    tblResult.Cell(i + 3, 7).Range.Text = col16[i];
                    tblResult.Cell(i + 3, 8).Range.Text = col17[i];
                    tblResult.Cell(i + 3, 9).Range.Text = "0";
                }
            }
            wordDoc.SaveAs2(ref filePath, ref fileFormat, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing);
            wordDoc.Close(ref nothing, ref nothing, ref nothing);
            wordApp.Quit();
            foreach (Process proc in Process.GetProcessesByName("WINWORD"))
            {
                proc.Kill();
            }
        }

        internal void OutputAccuracy(string p, double unitError, List<string> col0, List<string> col6, List<string> col7, List<string> col8, List<string> col9, List<string> col10, List<string> col11, List<string> col12, List<string> col13)
        {
            object nothing = Missing.Value;
            object fileFormat = WdSaveFormat.wdFormatDocument;
            object filePath = p;
            Microsoft.Office.Interop.Word._Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
            Microsoft.Office.Interop.Word._Document wordDoc = wordApp.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);
            wordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;
            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter("平差助手");
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
            wordApp.Selection.ParagraphFormat.LineSpacing = 15f;
            object count = col0.Count + 6;
            object wdLine = WdUnits.wdLine;
            for (int i = 0; i < 4; i++)
            {
                wordApp.Selection.MoveDown(ref wdLine, ref count, ref nothing);
                wordApp.Selection.TypeParagraph();
            }
            wordDoc.Paragraphs.First.Range.Text = "精度评定表";
            wordDoc.Paragraphs.First.Range.Font.Bold = 2;
            wordDoc.Paragraphs.First.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordDoc.Paragraphs[2].Range.Text = string.Format("{0,-50}{1,-50}", "工程名称：", "计算者：");
            wordDoc.Paragraphs[2].Range.Font.Bold = 0;
            wordDoc.Paragraphs[2].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordDoc.Paragraphs.Last.Range.Text = string.Format("{0,-50}{1,-50}", "校核者：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            Table tblAccuracy = wordDoc.Tables.Add(wordApp.Selection.Range, col0.Count + 3, 11, ref nothing, ref nothing);
            tblAccuracy.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleThickThinLargeGap;
            tblAccuracy.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            tblAccuracy.Columns[1].Width = 40f;
            tblAccuracy.Columns[2].Width = 50f;
            tblAccuracy.Columns[3].Width = 50f;
            tblAccuracy.Columns[4].Width = 50f;
            tblAccuracy.Columns[5].Width = 50f;
            tblAccuracy.Columns[6].Width = 50f;
            tblAccuracy.Columns[7].Width = 80f;
            tblAccuracy.Columns[8].Width = 50f;
            tblAccuracy.Columns[9].Width = 40f;
            tblAccuracy.Columns[10].Width = 60f;
            tblAccuracy.Columns[11].Width = 60f;
            tblAccuracy.Cell(1, 1).Range.Text = "点名";
            tblAccuracy.Cell(1, 1).Merge(tblAccuracy.Cell(2, 1));
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblAccuracy.Cell(1, 2).Range.Text = "点位中误差(m)";
            tblAccuracy.Cell(1, 2).Merge(tblAccuracy.Cell(1, 4));
            tblAccuracy.Cell(1, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblAccuracy.Cell(2, 2).Range.Text = "Mx";
            tblAccuracy.Cell(2, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblAccuracy.Cell(2, 3).Range.Text = "My";
            tblAccuracy.Cell(2, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblAccuracy.Cell(2, 4).Range.Text = "M";
            tblAccuracy.Cell(2, 4).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblAccuracy.Cell(1, 3).Range.Text = "误差椭圆";
            tblAccuracy.Cell(1, 3).Merge(tblAccuracy.Cell(1, 5));
            tblAccuracy.Cell(1, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblAccuracy.Cell(2, 5).Range.Text = "A(m)";
            tblAccuracy.Cell(2, 5).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblAccuracy.Cell(2, 6).Range.Text = "B(m)";
            tblAccuracy.Cell(2, 6).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblAccuracy.Cell(2, 7).Range.Text = "F(° ' \")";
            tblAccuracy.Cell(2, 7).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblAccuracy.Cell(1, 4).Range.Text = "高差中误差(m)";
            tblAccuracy.Cell(1, 4).Merge(tblAccuracy.Cell(2, 8));
            tblAccuracy.Cell(1, 4).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblAccuracy.Cell(1, 5).Range.Text = "至点";
            tblAccuracy.Cell(1, 5).Merge(tblAccuracy.Cell(2, 9));
            tblAccuracy.Cell(1, 5).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblAccuracy.Cell(1, 6).Range.Text = "方位角中误差(\")";
            tblAccuracy.Cell(1, 6).Merge(tblAccuracy.Cell(2, 10));
            tblAccuracy.Cell(1, 6).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblAccuracy.Cell(1, 7).Range.Text = "边长中误差(m)";
            tblAccuracy.Cell(1, 7).Merge(tblAccuracy.Cell(2, 11));
            tblAccuracy.Cell(1, 7).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblAccuracy.Cell(col0.Count + 3, 1).Range.Text = "备注";
            tblAccuracy.Cell(col0.Count + 3, 1).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblAccuracy.Cell(col0.Count + 3, 2).Range.Text = "单位权中误差 = " + unitError;
            tblAccuracy.Cell(col0.Count + 3, 2).Merge(tblAccuracy.Cell(col0.Count + 3, 11));
            tblAccuracy.Cell(col0.Count + 3, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            object moveUnit = WdUnits.wdLine;
            object moveExtend = WdMovementType.wdExtend;
            for (int i = 0; i < col0.Count; i++)
            {
                if (i == 0 || i == 1 || i == col0.Count - 2)
                {
                    tblAccuracy.Cell(i + 3, 1).Range.Text = col0[i];
                    tblAccuracy.Cell(i + 3, 2).Range.Text = col6[i];
                    tblAccuracy.Cell(i + 3, 3).Range.Text = col7[i];
                    tblAccuracy.Cell(i + 3, 4).Range.Text = col8[i];
                    tblAccuracy.Cell(i + 3, 5).Range.Text = " ";
                    tblAccuracy.Cell(i + 3, 6).Range.Text = " ";
                    tblAccuracy.Cell(i + 3, 7).Range.Text = " ";
                    tblAccuracy.Cell(i + 3, 8).Range.Text = "0";
                    tblAccuracy.Cell(i + 3, 9).Range.Text = col0[i + 1];
                    tblAccuracy.Cell(i + 3, 10).Range.Text = col10[i];
                    tblAccuracy.Cell(i + 3, 11).Range.Text = col9[i];
                }
                else if (i == col0.Count - 1)
                {
                    tblAccuracy.Cell(i + 3, 1).Range.Text = col0[i];
                    tblAccuracy.Cell(i + 3, 2).Range.Text = col6[i];
                    tblAccuracy.Cell(i + 3, 3).Range.Text = col7[i];
                    tblAccuracy.Cell(i + 3, 4).Range.Text = col8[i];
                    tblAccuracy.Cell(i + 3, 5).Range.Text = " ";
                    tblAccuracy.Cell(i + 3, 6).Range.Text = " ";
                    tblAccuracy.Cell(i + 3, 7).Range.Text = " ";
                    tblAccuracy.Cell(i + 3, 8).Range.Text = " ";
                    tblAccuracy.Cell(i + 3, 9).Range.Text = " ";
                    tblAccuracy.Cell(i + 3, 10).Range.Text = " ";
                    tblAccuracy.Cell(i + 3, 11).Range.Text = " ";
                }
                else
                {
                    tblAccuracy.Cell(i + 3, 1).Range.Text = col0[i];
                    tblAccuracy.Cell(i + 3, 2).Range.Text = col6[i];
                    tblAccuracy.Cell(i + 3, 3).Range.Text = col7[i];
                    tblAccuracy.Cell(i + 3, 4).Range.Text = col8[i];
                    tblAccuracy.Cell(i + 3, 5).Range.Text = col11[i - 2];
                    tblAccuracy.Cell(i + 3, 6).Range.Text = col12[i - 2];
                    tblAccuracy.Cell(i + 3, 7).Range.Text = col13[i - 2];
                    tblAccuracy.Cell(i + 3, 8).Range.Text = "0";
                    tblAccuracy.Cell(i + 3, 9).Range.Text = col0[i + 1];
                    tblAccuracy.Cell(i + 3, 10).Range.Text = col10[i];
                    tblAccuracy.Cell(i + 3, 11).Range.Text = col9[i];
                }
            }
            wordDoc.SaveAs2(ref filePath, ref fileFormat, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing);
            wordDoc.Close(ref nothing, ref nothing, ref nothing);
            wordApp.Quit();
            foreach (Process proc in Process.GetProcessesByName("WINWORD"))
            {
                proc.Kill();
            }
        }

        internal void OutputBLToXY(string p, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5)
        {
            object nothing = Missing.Value;
            object fileFormat = WdSaveFormat.wdFormatDocument;
            object filePath = p;
            Microsoft.Office.Interop.Word._Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
            Microsoft.Office.Interop.Word._Document wordDoc = wordApp.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);
            wordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;
            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter("平差助手");
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
            wordApp.Selection.ParagraphFormat.LineSpacing = 15f;
            object count = col1.Count + 5;
            object wdLine = WdUnits.wdLine;
            for (int i = 0; i < 4; i++)
            {
                wordApp.Selection.MoveDown(ref wdLine, ref count, ref nothing);
                wordApp.Selection.TypeParagraph();
            }
            wordDoc.Paragraphs.First.Range.Text = "高斯投影正算";
            wordDoc.Paragraphs.First.Range.Font.Bold = 2;
            wordDoc.Paragraphs.First.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordDoc.Paragraphs[2].Range.Text = string.Format("{0,-37}{1,-37}", "工程名称：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            wordDoc.Paragraphs[2].Range.Font.Bold = 0;
            wordDoc.Paragraphs[2].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordDoc.Paragraphs.Last.Range.Text = string.Format("{0,-25}{1,-25}{2,-25}", "计算：", "复核：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));

            Table tblBLToXY = wordDoc.Tables.Add(wordApp.Selection.Range, col1.Count + 2, 5, ref nothing, ref nothing);
            tblBLToXY.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleThickThinLargeGap;
            tblBLToXY.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            tblBLToXY.Columns[1].Width = 60f;
            tblBLToXY.Columns[2].Width = 110f;
            tblBLToXY.Columns[3].Width = 110f;
            tblBLToXY.Columns[4].Width = 100f;
            tblBLToXY.Columns[5].Width = 100f;
            tblBLToXY.Cell(1, 1).Range.Text = "点名";
            tblBLToXY.Cell(1, 1).Merge(tblBLToXY.Cell(2, 1));
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblBLToXY.Cell(1, 2).Range.Text = "大地坐标(° '  \")";
            tblBLToXY.Cell(1, 2).Merge(tblBLToXY.Cell(1, 3));
            tblBLToXY.Cell(1, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblBLToXY.Cell(2, 2).Range.Text = "纬度";
            tblBLToXY.Cell(2, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblBLToXY.Cell(2, 3).Range.Text = "经度";
            tblBLToXY.Cell(2, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblBLToXY.Cell(1, 3).Range.Text = "高斯坐标(m)";
            tblBLToXY.Cell(1, 3).Merge(tblBLToXY.Cell(1, 4));
            tblBLToXY.Cell(1, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblBLToXY.Cell(2, 4).Range.Text = "X";
            tblBLToXY.Cell(2, 4).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblBLToXY.Cell(2, 5).Range.Text = "Y";
            tblBLToXY.Cell(2, 5).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            object moveUnit = WdUnits.wdLine;
            object moveExtend = WdMovementType.wdExtend;
            for (int i = 0; i < col0.Count; i++)
            {
                tblBLToXY.Cell(i + 3, 1).Range.Text = col0[i];
                tblBLToXY.Cell(i + 3, 2).Range.Text = col1[i];
                tblBLToXY.Cell(i + 3, 3).Range.Text = col2[i];
                tblBLToXY.Cell(i + 3, 4).Range.Text = col4[i];
                tblBLToXY.Cell(i + 3, 5).Range.Text = col5[i];
            }

            wordDoc.SaveAs2(ref filePath, ref fileFormat, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing);
            wordDoc.Close(ref nothing, ref nothing, ref nothing);
            wordApp.Quit();
            foreach (Process proc in Process.GetProcessesByName("WINWORD"))
            {
                proc.Kill();
            }
        }

        internal void OutputXYToBL(string p, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5)
        {
            object nothing = Missing.Value;
            object fileFormat = WdSaveFormat.wdFormatDocument;
            object filePath = p;
            Microsoft.Office.Interop.Word._Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
            Microsoft.Office.Interop.Word._Document wordDoc = wordApp.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);
            wordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;
            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter("平差助手");
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
            wordApp.Selection.ParagraphFormat.LineSpacing = 15f;
            object count = col1.Count + 5;
            object wdLine = WdUnits.wdLine;
            for (int i = 0; i < 4; i++)
            {
                wordApp.Selection.MoveDown(ref wdLine, ref count, ref nothing);
                wordApp.Selection.TypeParagraph();
            }
            wordDoc.Paragraphs.First.Range.Text = "高斯投影反算";
            wordDoc.Paragraphs.First.Range.Font.Bold = 2;
            wordDoc.Paragraphs.First.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordDoc.Paragraphs[2].Range.Text = string.Format("{0,-37}{1,-37}", "工程名称：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            wordDoc.Paragraphs[2].Range.Font.Bold = 0;
            wordDoc.Paragraphs[2].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordDoc.Paragraphs.Last.Range.Text = string.Format("{0,-25}{1,-25}{2,-25}", "计算：", "复核：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));

            Table tblXYToBL = wordDoc.Tables.Add(wordApp.Selection.Range, col1.Count + 2, 5, ref nothing, ref nothing);
            tblXYToBL.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleThickThinLargeGap;
            tblXYToBL.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            tblXYToBL.Columns[1].Width = 60f;
            tblXYToBL.Columns[2].Width = 100f;
            tblXYToBL.Columns[3].Width = 100f;
            tblXYToBL.Columns[4].Width = 110f;
            tblXYToBL.Columns[5].Width = 110f;
            tblXYToBL.Cell(1, 1).Range.Text = "点名";
            tblXYToBL.Cell(1, 1).Merge(tblXYToBL.Cell(2, 1));
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblXYToBL.Cell(1, 2).Range.Text = "高斯坐标(m)";
            tblXYToBL.Cell(1, 2).Merge(tblXYToBL.Cell(1, 3));
            tblXYToBL.Cell(1, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblXYToBL.Cell(2, 2).Range.Text = "X";
            tblXYToBL.Cell(2, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblXYToBL.Cell(2, 3).Range.Text = "Y";
            tblXYToBL.Cell(2, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblXYToBL.Cell(1, 3).Range.Text = "大地坐标(° '  \")";
            tblXYToBL.Cell(1, 3).Merge(tblXYToBL.Cell(1, 4));
            tblXYToBL.Cell(1, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblXYToBL.Cell(2, 4).Range.Text = "经度";
            tblXYToBL.Cell(2, 4).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblXYToBL.Cell(2, 5).Range.Text = "纬度";
            tblXYToBL.Cell(2, 5).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            object moveUnit = WdUnits.wdLine;
            object moveExtend = WdMovementType.wdExtend;
            for (int i = 0; i < col0.Count; i++)
            {
                tblXYToBL.Cell(i + 3, 1).Range.Text = col0[i];
                tblXYToBL.Cell(i + 3, 2).Range.Text = col1[i];
                tblXYToBL.Cell(i + 3, 3).Range.Text = col2[i];
                tblXYToBL.Cell(i + 3, 4).Range.Text = col4[i];
                tblXYToBL.Cell(i + 3, 5).Range.Text = col5[i];
            }

            wordDoc.SaveAs2(ref filePath, ref fileFormat, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing);
            wordDoc.Close(ref nothing, ref nothing, ref nothing);
            wordApp.Quit();
            foreach (Process proc in Process.GetProcessesByName("WINWORD"))
            {
                proc.Kill();
            }
        }

        internal void OutputXYToXY(string p, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5)
        {
            object nothing = Missing.Value;
            object fileFormat = WdSaveFormat.wdFormatDocument;
            object filePath = p;
            Microsoft.Office.Interop.Word._Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
            Microsoft.Office.Interop.Word._Document wordDoc = wordApp.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);
            wordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;
            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter("平差助手");
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
            wordApp.Selection.ParagraphFormat.LineSpacing = 15f;
            object count = col1.Count + 5;
            object wdLine = WdUnits.wdLine;
            for (int i = 0; i < 4; i++)
            {
                wordApp.Selection.MoveDown(ref wdLine, ref count, ref nothing);
                wordApp.Selection.TypeParagraph();
            }
            wordDoc.Paragraphs.First.Range.Text = "坐标换带计算";
            wordDoc.Paragraphs.First.Range.Font.Bold = 2;
            wordDoc.Paragraphs.First.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordDoc.Paragraphs[2].Range.Text = string.Format("{0,-37}{1,-37}", "工程名称：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));
            wordDoc.Paragraphs[2].Range.Font.Bold = 0;
            wordDoc.Paragraphs[2].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordDoc.Paragraphs.Last.Range.Text = string.Format("{0,-25}{1,-25}{2,-25}", "计算：", "复核：", "日期：" + DateTime.Now.ToString("yyyy-MM-dd"));

            Table tblXYToBL = wordDoc.Tables.Add(wordApp.Selection.Range, col1.Count + 2, 5, ref nothing, ref nothing);
            tblXYToBL.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleThickThinLargeGap;
            tblXYToBL.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            tblXYToBL.Columns[1].Width = 60f;
            tblXYToBL.Columns[2].Width = 100f;
            tblXYToBL.Columns[3].Width = 100f;
            tblXYToBL.Columns[4].Width = 100f;
            tblXYToBL.Columns[5].Width = 100f;
            tblXYToBL.Cell(1, 1).Range.Text = "点名";
            tblXYToBL.Cell(1, 1).Merge(tblXYToBL.Cell(2, 1));
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblXYToBL.Cell(1, 2).Range.Text = "转换前坐标(m)";
            tblXYToBL.Cell(1, 2).Merge(tblXYToBL.Cell(1, 3));
            tblXYToBL.Cell(1, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblXYToBL.Cell(2, 2).Range.Text = "X";
            tblXYToBL.Cell(2, 2).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblXYToBL.Cell(2, 3).Range.Text = "Y";
            tblXYToBL.Cell(2, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblXYToBL.Cell(1, 3).Range.Text = "转换后坐标(m)";
            tblXYToBL.Cell(1, 3).Merge(tblXYToBL.Cell(1, 4));
            tblXYToBL.Cell(1, 3).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblXYToBL.Cell(2, 4).Range.Text = "X";
            tblXYToBL.Cell(2, 4).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblXYToBL.Cell(2, 5).Range.Text = "Y";
            tblXYToBL.Cell(2, 5).Select();
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            object moveUnit = WdUnits.wdLine;
            object moveExtend = WdMovementType.wdExtend;
            for (int i = 0; i < col0.Count; i++)
            {
                tblXYToBL.Cell(i + 3, 1).Range.Text = col0[i];
                tblXYToBL.Cell(i + 3, 2).Range.Text = col1[i];
                tblXYToBL.Cell(i + 3, 3).Range.Text = col2[i];
                tblXYToBL.Cell(i + 3, 4).Range.Text = col4[i];
                tblXYToBL.Cell(i + 3, 5).Range.Text = col5[i];
            }

            wordDoc.SaveAs2(ref filePath, ref fileFormat, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing);
            wordDoc.Close(ref nothing, ref nothing, ref nothing);
            wordApp.Quit();
            foreach (Process proc in Process.GetProcessesByName("WINWORD"))
            {
                proc.Kill();
            }
        }
    }
}
