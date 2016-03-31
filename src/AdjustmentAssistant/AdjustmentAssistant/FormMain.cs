using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Text;

namespace AdjustmentAssistant
{
    public partial class FormMain : Form
    {
        #region
        private List<string> col0 = new List<string>();
        private List<string> col1 = new List<string>();
        private List<string> col2 = new List<string>();
        private List<string> col3 = new List<string>();
        private List<string> col4 = new List<string>();
        private List<string> col5 = new List<string>();
        private List<string> col6 = new List<string>();
        private List<string> col7 = new List<string>();
        private List<string> col8 = new List<string>();
        private List<string> col9 = new List<string>();
        private List<string> col10 = new List<string>();
        private List<string> col11 = new List<string>();
        private List<string> col12 = new List<string>();
        private List<string> col13 = new List<string>();
        private List<string> col14 = new List<string>();
        private List<string> col15 = new List<string>();
        private List<string> col16 = new List<string>();
        private List<string> col17 = new List<string>();
        private List<string> col18 = new List<string>();
        private DataType.DataType dataType = new DataType.DataType();
        private DataType.Data approximateDataType = new DataType.Data();
        private DataType.LeftOrRight lorR = new DataType.LeftOrRight();
        private DataType.LevelingWeight LvlWet = new DataType.LevelingWeight();
        private int dataCount;
        private int backCount;
        private double coordinateCloseError;
        private double angleCloseError;
        private double k;
        private string strTitle;
        private string strProjectName;
        private string strInstrument;
        private string strWeather;
        private string strObserver;
        private string strRecorder;
        private string strDate;
        private string strCalculate;
        private string strAssessment;
        private TableLayoutPanel tableLayoutPanel;
        private bool isCalculate;
        private string filePath;
        private int tableType;
        private Bitmap b;
        private double[] accuracy = new double[6];
        private double unitError;
        private int gKNo;
        private int gKOrCoord;
        private int inputMidLon;
        private int outputMidLon;
        private string isRegedit = "试用版";
        private string cdKey;
        private bool regeditStatus = false;
        #endregion

        public FormMain()
        {
            InitializeComponent();
        }

        #region 工具箱的动态变化
        private void InitialToolBox()
        {
            pnlRecord.Height = 38;
            pnlAdjust.Location = new Point(0, 38);
            pnlAdjust.Height = 38;
            pnlTool.Location = new Point(3, 76);
            pnlTool.Height = 38;
        }

        private void btnRecord_Click(object sender, EventArgs e)
        {
            InitialToolBox();
            pnlRecord.Height = splitContainer1.Panel1.Height - 80;
            treeRecord.Height = splitContainer1.Panel1.Height - 120;
        }

        private void btnAdjust_Click(object sender, EventArgs e)
        {
            InitialToolBox();
            pnlAdjust.Height = splitContainer1.Panel1.Height - 80;
            treeAdjust.Height = splitContainer1.Panel1.Height - 120;
        }

        private void btnTool_Click(object sender, EventArgs e)
        {
            InitialToolBox();
            pnlTool.Height = splitContainer1.Panel1.Height - 80;
            treeTool.Height = splitContainer1.Panel1.Height - 120;
        }
        #endregion

        private void FormMain_Load(object sender, EventArgs e)
        {
            IsRegedit();
        }

        private void IsRegedit()
        {
            string keyPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Set.ini");
            System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create();
            byte[] md5Buffer = md5.ComputeHash(System.Text.Encoding.Default.GetBytes(GetCDKey.GetCpuID()));
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < md5Buffer.Length; i += 2)
            {
                if ((i != 14) && (i / 2) % 2 == 1)
                {
                    sb.Append(md5Buffer[i].ToString("X2") + "-");
                }
                else
                {
                    sb.Append(md5Buffer[i].ToString("X2"));
                }
            }
            cdKey = sb.ToString();
            using (StreamReader sr = new StreamReader("Set.ini"))
            {
                if (sr.ReadLine() == cdKey)
                {
                    regeditStatus = true;
                    this.menuHelpRegedit.Text = "已注册(&R)";
                    this.menuHelpRegedit.Enabled = false;
                    isRegedit = "已注册给用户 " + Environment.UserName;
                }
            }
        }

        private void menuFileOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDlg = new OpenFileDialog();
            openDlg.Filter = "全站仪数据(*.dat)|*.dat|文本文件(*.txt)|*.txt";
            openDlg.Title = "打开测量数据";
            if (openDlg.ShowDialog() == DialogResult.OK)
            {
                ClearAll();
                tableType = 0;
                pnlResult.Controls.Clear();
                dataType = ReadData.OpenFile(openDlg.FileName, ref col0, ref col1, ref col2, ref col3);
                isCalculate = false;
            }
        }

        private void treeRecord_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (treeRecord.SelectedNode.Index == 0)
            {
                if (approximateDataType == DataType.Data.ConnectingTraverse || approximateDataType == DataType.Data.OpenTraverse)
                {
                    if (isCalculate == false)
                    {
                        tableLayoutPanel = new TableLayoutPanel();
                        tableLayoutPanel.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(tableLayoutPanel, true, null);
                        TraverseRecordTable.DrawTable(pnlResult, dataCount, ref tableLayoutPanel);
                        InputTraverseRecordTable.GetData(pnlResult, dataCount, backCount, approximateDataType, ref tableLayoutPanel);
                        dataType = DataType.DataType.HandMade;
                        tableType = 1;
                    }
                }
            }
            else if (treeRecord.SelectedNode.Index == 1)
            {
                if (approximateDataType == DataType.Data.ConnectingTraverse || approximateDataType == DataType.Data.OpenTraverse)
                {
                    if (isCalculate == false)
                    {
                        tableLayoutPanel = new TableLayoutPanel();
                        tableLayoutPanel.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(tableLayoutPanel, true, null);
                        LevelingRecordTable.DrawTable(pnlResult, dataCount, ref tableLayoutPanel);
                        InputLevelingRecordTable.GetData(pnlResult, dataCount, backCount, approximateDataType, ref tableLayoutPanel);
                        dataType = DataType.DataType.HandMade;
                        tableType = 2;
                    }
                }
            }
            else if (treeRecord.SelectedNode.Index == 2)
            {

            }
        }

        private void treeAdjust_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (treeAdjust.SelectedNode.Index == 0)
            {
                if (dataType == DataType.DataType.TotalStation && isCalculate == false)
                {
                    b = null;
                    pnlResult.BackgroundImage = null;
                    TraverseAdjustment traverseAdjustTable = new TraverseAdjustment();
                    tableLayoutPanel = new TableLayoutPanel();
                    tableLayoutPanel.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(tableLayoutPanel, true, null);
                    traverseAdjustTable.DrawTable(pnlResult, col0, col1, col2, col3, tableLayoutPanel);//此处col3为高程值，为预留接口
                    tableType = 4;
                }
                else
                {
                    b = null;
                    pnlResult.BackgroundImage = null;
                    TraverseAdjustment traverseAdjustTable = new TraverseAdjustment();
                    tableLayoutPanel = new TableLayoutPanel();
                    tableLayoutPanel.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(tableLayoutPanel, true, null);
                    traverseAdjustTable.DrawTable(pnlResult, col0, col1, col2, col4, col5, col8, col14, col15, tableLayoutPanel);//此处col3为高程值，为预留接口
                }
            }
            else if (treeAdjust.SelectedNode.Index == 1)
            {

            }
            else if (treeAdjust.SelectedNode.Index == 2)
            {
                if (dataType == DataType.DataType.TotalStation && isCalculate == true)
                {
                    pnlResult.BackgroundImage = null;
                    ControlPointTable controlPointTable = new ControlPointTable();
                    tableLayoutPanel = new TableLayoutPanel();
                    tableLayoutPanel.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(tableLayoutPanel, true, null);
                    controlPointTable.DrawTable(filePath, pnlResult, tableLayoutPanel, col0, col14, col15);
                    tableType = 6;
                }
            }
            else if (treeAdjust.SelectedNode.Index == 3)
            {
                if (dataType == DataType.DataType.TotalStation && isCalculate == true)
                {
                    pnlResult.BackgroundImage = null;
                    ResultTable resultTable = new ResultTable();
                    tableLayoutPanel = new TableLayoutPanel();
                    tableLayoutPanel.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(tableLayoutPanel, true, null);
                    resultTable.DrawTable(filePath, pnlResult, tableLayoutPanel, col0, col14, col15, col16, col17, col18);
                    tableType = 7;
                }
            }
            else if (treeAdjust.SelectedNode.Index == 4)
            {
                if (dataType == DataType.DataType.TotalStation && isCalculate == true)
                {
                    pnlResult.BackgroundImage = null;
                    AccuracyTable accuracyTable = new AccuracyTable();
                    tableLayoutPanel = new TableLayoutPanel();
                    tableLayoutPanel.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(tableLayoutPanel, true, null);
                    accuracyTable.DrawTable(filePath, pnlResult, tableLayoutPanel, unitError, col0, col6, col7, col8, col9, col10, col11, col12, col13);
                    tableType = 8;
                }
            }
            else if (treeAdjust.SelectedNode.Index == 5)
            {
                if (dataType == DataType.DataType.TotalStation && isCalculate == true)
                {
                    pnlResult.BackgroundImage = null;
                    DrawGraphics drawGraphic = new DrawGraphics();
                    b = drawGraphic.DrawPoint(pnlResult, col0, col14, col15);
                    tableType = 9;
                }
            }
            else if (treeAdjust.SelectedNode.Index == 6)
            {
                if (dataType == DataType.DataType.TotalStation && isCalculate == true)
                {
                    pnlResult.BackgroundImage = null;
                    DrawGraphics drawGraphic = new DrawGraphics();
                    b = drawGraphic.DrawLine(pnlResult, col0, col11, col12, col13, col14, col15);
                    tableType = 10;
                }
            }
        }

        private void treeTool_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (treeTool.SelectedNode.Index == 0)
            {
                if (dataType == DataType.DataType.Geodetic && isCalculate == false)
                {
                    BLToXYTable bLtoXY = new BLToXYTable();
                    tableLayoutPanel = new TableLayoutPanel();
                    tableLayoutPanel.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(tableLayoutPanel, true, null);
                    bLtoXY.DrawTable(pnlResult, col0.Count, ref tableLayoutPanel);
                    InputBLToXY.GetData(col0, col1, col2, ref tableLayoutPanel);
                    tableType = 11;
                }
            }
            else if (treeTool.SelectedNode.Index == 1)
            {
                if (dataType == DataType.DataType.Gauss && isCalculate == false)
                {
                    XYToBLTable xYtoBL = new XYToBLTable();
                    tableLayoutPanel = new TableLayoutPanel();
                    tableLayoutPanel.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(tableLayoutPanel, true, null);
                    xYtoBL.DrawTable(pnlResult, col0.Count, ref tableLayoutPanel);
                    InputXYToBL.GetData(col0, col1, col2, ref tableLayoutPanel);
                    tableType = 12;
                }
            }
            else if (treeTool.SelectedNode.Index == 2)
            {
                if (dataType == DataType.DataType.Gauss && isCalculate == false)
                {
                    XYToXYTable xYtoXY = new XYToXYTable();
                    tableLayoutPanel = new TableLayoutPanel();
                    tableLayoutPanel.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(tableLayoutPanel, true, null);
                    xYtoXY.DrawTable(pnlResult, col0.Count, ref tableLayoutPanel);
                    InputXYToXY.GetData(col0, col1, col2, ref tableLayoutPanel);
                    tableType = 13;
                }
            }
        }

        private void tsbtnFormat_Click(object sender, EventArgs e)
        {
            if (isCalculate == true)
            {
                isCalculate = false;
            }
            if (dataType == DataType.DataType.ApproximateAdjustment || dataType == DataType.DataType.HandMade)
            {
                FormFormat frm = new FormFormat();
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    dataCount = frm.DataCount;
                    backCount = frm.BackCount;
                    approximateDataType = frm.TraverseDataType;
                    lorR = frm.LorR;
                    LvlWet = frm.LvlWet;
                }
            }
            else if (dataType == DataType.DataType.TotalStation)
            {
                FormAccuracy frm = new FormAccuracy();
                frm.lblPointName = col0[col0.Count - 2];
                frm.pointX = Convert.ToDouble(col1[col1.Count - 2]);
                frm.pointY = Convert.ToDouble(col2[col2.Count - 2]);
                frm.Init();
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    accuracy[0] = frm.direction;
                    accuracy[1] = frm.side;
                    accuracy[2] = frm.sidePercent;
                    accuracy[3] = frm.angle;
                    accuracy[4] = frm.pointX;
                    accuracy[5] = frm.pointY;
                }
            }
            else if (dataType == DataType.DataType.Geodetic || dataType == DataType.DataType.Gauss)
            {
                FormCoordinate frm = new FormCoordinate();
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    gKNo = frm.gKNo;//高斯投影坐标分带
                    gKOrCoord = frm.gKOrCoord;//高斯投影正反算还是坐标换带转换
                    inputMidLon = frm.inputMidLon;
                    outputMidLon = frm.outputMidLon;
                }
            }
        }

        private void menuOperateCalculate_Click(object sender, EventArgs e)
        {
            if (isCalculate == true)
            {
                return;
            }
            if (tableType == 1)
            {
                ClearAll();
                filePath = Calculate.TraverseRecord(dataCount, backCount, approximateDataType, lorR, ref tableLayoutPanel);
                OutputTraverseRecordTable.GetData(filePath, backCount, tableLayoutPanel, ref coordinateCloseError, ref angleCloseError, ref k, ref col0, ref col1, ref col2, ref col3, ref col4, ref col5, ref col6, ref col7, ref col8, ref col9, ref col10);
                isCalculate = true;
            }
            else if (tableType == 2)
            {
                ClearAll();
                filePath = Calculate.LevelingRecord(dataCount, backCount, approximateDataType, lorR, ref tableLayoutPanel);
                OutputLevelingRecordTable.GetData(filePath, backCount, tableLayoutPanel, ref col0, ref col1, ref col2, ref col3, ref col4, ref col5, ref col6);
                isCalculate = true;
            }
            else if (tableType == 4)
            {
                GetParameter();
                filePath = Calculate.ParameterAdjustment(col0, col1, col2, ref col4, ref col5, accuracy, approximateDataType, lorR, ref tableLayoutPanel);
                OutputParameterAdjustment.GetData(filePath, tableLayoutPanel, ref unitError, col0, col1, col2, col4, col5, ref col6, ref col7, ref col8, ref col9, ref col10, ref col11, ref col12, ref col13, ref col14, ref col15, ref col16, ref col17, ref col18);
                isCalculate = true;
            }
            else if (tableType == 11)
            {
                GetParameter();
                Adjustment.ConvertCoordinate.GetBLToXYTable(gKNo, col0, col1, col2, ref col4, ref col5);
                OutputBLToXY.GetData(ref tableLayoutPanel, col0, col1, col2, col4, col5);
                isCalculate = true;
            }
            else if (tableType == 12)
            {
                GetParameter();
                Adjustment.ConvertCoordinate.GetXYToBLTable(inputMidLon, col0, col1, col2, ref col4, ref col5);
                OutputXYToBL.GetData(ref tableLayoutPanel, col0, col1, col2, col4, col5);
                isCalculate = true;
            }
            else if (tableType == 13)
            {
                GetParameter();
                Adjustment.ConvertCoordinate.GetXYToXYTable(gKNo, inputMidLon, outputMidLon, col0, col1, col2, ref col4, ref col5);
                OutputXYToXY.GetData(ref tableLayoutPanel, col0, col1, col2, col4, col5);
                isCalculate = true;
            }
        }

        private void tsbtnClean_Click(object sender, EventArgs e)
        {
            ClearAll();
            tableType = 0;
            pnlResult.Controls.Clear();
        }

        private void ClearAll()
        {
            col0.Clear();
            col1.Clear();
            col2.Clear();
            col3.Clear();
            GetParameter();
        }

        private void GetParameter()
        {
            col4.Clear();
            col5.Clear();
            col6.Clear();
            col7.Clear();
            col8.Clear();
            col9.Clear();
            col10.Clear();
            col11.Clear();
            col12.Clear();
            col13.Clear();
            col14.Clear();
            col15.Clear();
            col16.Clear();
            col17.Clear();
            col18.Clear();
            isCalculate = false;
        }

        private void menuHelpRegedit_Click(object sender, EventArgs e)
        {
            FormRegedit frm = new FormRegedit(cdKey);
            frm.ShowDialog();
        }

        private void menuHelpAbout_Click(object sender, EventArgs e)
        {
            FormAbout frm = new FormAbout(isRegedit);
            frm.ShowDialog();
        }

        private void tsbtnTxt_Click(object sender, EventArgs e)
        {
            if (regeditStatus == false)
            {
                MessageBox.Show("试用版此功能受到限制");
                return;
            }
            if (isCalculate == false)
            {
                MessageBox.Show("没有得到计算结果");
                return;
            }
            ToText();
        }

        private void ToText()
        {
            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.Filter = "文本文件(*.txt)|*.txt";
            if (saveDlg.ShowDialog() == DialogResult.OK)
            {
                OutputText outputTxt = new OutputText();
                if (tableType == 1)
                {
                    outputTxt.OutputTraverse(saveDlg.FileName, approximateDataType, backCount, angleCloseError, coordinateCloseError, k, col0, col1, col2, col3, col4, col5, col6, col7, col8, col9, col10);
                }
                else if (tableType == 2)
                {
                    outputTxt.OutputLevelAngle(saveDlg.FileName, approximateDataType, backCount, angleCloseError, col0, col1, col2, col3, col4, col5, col6);
                }
                else if (tableType == 4)
                {
                    outputTxt.OutputPlane(saveDlg.FileName, col0, col1, col2, col4, col5, col8, col14, col15);
                }
                else if (tableType == 6)
                {
                    outputTxt.OutputPoint(saveDlg.FileName, col0, col14, col15);
                }
                else if (tableType == 7)
                {
                    outputTxt.OutputResult(saveDlg.FileName, col0, col14, col15, col18, col16, col17);
                }
                else if (tableType == 8)
                {
                    outputTxt.OutputAccuracy(saveDlg.FileName, unitError, col0, col6, col7, col8, col9, col10, col11, col12, col13);
                }
                else if (tableType == 11)
                {
                    outputTxt.OutputBLToXY(saveDlg.FileName, col0, col1, col2, col4, col5);
                }
                else if (tableType == 12)
                {
                    outputTxt.OutputXYToBL(saveDlg.FileName, col0, col1, col2, col4, col5);
                }
                else if (tableType == 13)
                {
                    outputTxt.OutputXYToXY(saveDlg.FileName, col0, col1, col2, col4, col5);
                }
            }
        }

        private void tsbtnDoc_Click(object sender, EventArgs e)
        {
            if (regeditStatus == false)
            {
                MessageBox.Show("试用版此功能受到限制");
                return;
            }
            if (isCalculate == false)
            {
                MessageBox.Show("没有得到计算结果");
                return;
            }
            ToWord();
        }

        private void ToWord()
        {
            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.Filter = "Word文档(*.doc)|*.doc";
            if (saveDlg.ShowDialog() == DialogResult.OK)
            {
                OutputWord outputDoc = new OutputWord();
                if (tableType == 1)
                {
                    outputDoc.OutputTraverse(saveDlg.FileName, approximateDataType, backCount, angleCloseError, coordinateCloseError, k, col0, col1, col2, col3, col4, col5, col6, col7, col8, col9, col10);
                }
                else if (tableType == 2)
                {
                    outputDoc.OutputLevelAngle(saveDlg.FileName, approximateDataType, backCount, angleCloseError, col0, col1, col2, col3, col4, col5, col6);
                }
                else if (tableType == 4)
                {
                    outputDoc.OutputPlane(saveDlg.FileName, col0, col1, col2, col4, col5, col8, col14, col15);
                }
                else if (tableType == 6)
                {
                    outputDoc.OutputPoint(saveDlg.FileName, col0, col14, col15);
                }
                else if (tableType == 7)
                {
                    outputDoc.OutputResult(saveDlg.FileName, col0, col14, col15, col16, col17, col18);
                }
                else if (tableType == 8)
                {
                    outputDoc.OutputAccuracy(saveDlg.FileName, unitError, col0, col6, col7, col8, col9, col10, col11, col12, col13);
                }
                else if (tableType == 11)
                {
                    outputDoc.OutputBLToXY(saveDlg.FileName, col0, col1, col2, col4, col5);
                }
                else if (tableType == 12)
                {
                    outputDoc.OutputXYToBL(saveDlg.FileName, col0, col1, col2, col4, col5);
                }
                else if (tableType == 13)
                {
                    outputDoc.OutputXYToXY(saveDlg.FileName, col0, col1, col2, col4, col5);
                }
            }
        }

        private void tsbtnXls_Click(object sender, EventArgs e)
        {
            if (regeditStatus == false)
            {
                MessageBox.Show("试用版此功能受到限制");
                return;
            }
            if (isCalculate == false)
            {
                MessageBox.Show("没有得到计算结果");
                return;
            }
            ToExcel();
        }

        private void ToExcel()
        {
            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.Filter = "Excel文档(*.xls)|*.xls";
            if (saveDlg.ShowDialog() == DialogResult.OK)
            {
                OutputExcel outputXls = new OutputExcel();
                if (tableType == 1)
                {
                    outputXls.OutputTraverse(saveDlg.FileName, approximateDataType, backCount, angleCloseError, coordinateCloseError, k, col0, col1, col2, col3, col4, col5, col6, col7, col8, col9, col10);
                }
                else if (tableType == 2)
                {
                    outputXls.OutputLevelAngle(saveDlg.FileName, approximateDataType, backCount, angleCloseError, col0, col1, col2, col3, col4, col5, col6);
                }
                else if (tableType == 4)
                {
                    outputXls.OutputPlane(saveDlg.FileName, col0, col1, col2, col4, col5, col8, col14, col15);
                }
                else if (tableType == 6)
                {
                    outputXls.OutputPoint(saveDlg.FileName, col0, col14, col15);
                }
                else if (tableType == 7)
                {
                    outputXls.OutputResult(saveDlg.FileName, col0, col14, col15, col16, col17, col18);
                }
                else if (tableType == 8)
                {
                    outputXls.OutputAccuracy(saveDlg.FileName, unitError, col0, col6, col7, col8, col9, col10, col11, col12, col13);
                }
                else if (tableType == 11)
                {
                    outputXls.OutputBLToXY(saveDlg.FileName, col0, col1, col2, col9, col10);
                }
                else if (tableType == 12)
                {
                    outputXls.OutputXYToBL(saveDlg.FileName, col0, col1, col2, col9, col10);
                }
                else if (tableType == 13)
                {
                    outputXls.OutputXYToXY(saveDlg.FileName, col0, col1, col2, col9, col10);
                }
            }
        }

        private void tsbtnDxf_Click(object sender, EventArgs e)
        {
            if (regeditStatus == false)
            {
                MessageBox.Show("试用版此功能受到限制");
                return;
            }
            if (isCalculate == false)
            {
                MessageBox.Show("没有得到计算结果");
                return;
            }
            ToDxf();
        }

        private void ToDxf()
        {
            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.Filter = "Dxf文件(.dxf)|*.dxf";
            if (saveDlg.ShowDialog() == DialogResult.OK)
            {
                OutputDxf outputDxf = new OutputDxf();
                if (tableType == 1)
                {
                    outputDxf.OutputTraverse(saveDlg.FileName, col0, col9, col10);
                }
                else if (tableType == 2)
                {
                    outputDxf.OutputTraverse(saveDlg.FileName, col0, col5, col6);
                }
                else if (tableType == 4)//此处保存的是观测记录！不是平差结果
                {
                    outputDxf.OutputTraverse(saveDlg.FileName, col0, col1, col2);
                }
                else if (tableType == 6 || tableType == 7 || tableType == 8 || tableType == 10)
                {
                    outputDxf.OutputTraverse(saveDlg.FileName, col0, col14, col15);
                }
                else if (tableType == 9)
                {
                    outputDxf.OutputPoint(saveDlg.FileName, col0, col14, col15);
                }
                else if (tableType == 11 || tableType == 13)
                {
                    outputDxf.OutputPoint(saveDlg.FileName, col0, col4, col5);
                }
                else if (tableType == 12)
                {
                    outputDxf.OutputPoint(saveDlg.FileName, col0, col1, col2);
                }
            }
        }

        private void tsbtnJpg_Click(object sender, EventArgs e)
        {
            if (regeditStatus == false)
            {
                MessageBox.Show("试用版此功能受到限制");
                return;
            }
            if (isCalculate == false)
            {
                MessageBox.Show("没有得到计算结果");
                return;
            }
            if (tableType == 9 || tableType == 10)
            {
                SaveFileDialog saveDlg = new SaveFileDialog();
                saveDlg.Filter = "保存为图片|*.jpg";
                if (saveDlg.ShowDialog() == DialogResult.OK)
                {
                    b.Save(saveDlg.FileName);
                }
            }
        }

        private void menuFileQuit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void menuHelpHelper_Click(object sender, EventArgs e)
        {
            Help.ShowHelp(new Control(), "Help.chm");
        }
    }
}
