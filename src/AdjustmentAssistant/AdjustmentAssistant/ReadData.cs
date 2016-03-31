using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace AdjustmentAssistant
{
    public static class ReadData
    {
        /// <summary>
        /// 用于读取文件
        /// </summary>
        /// <param name="filePath">读取的文件路径</param>
        /// <param name="col0">点名</param>
        /// <param name="col1">X或B</param>
        /// <param name="col2">Y或L</param>
        /// <param name="col3">Z</param>
        public static DataType.DataType OpenFile(string filePath, ref List<string> col0, ref List<string> col1, ref List<string> col2, ref List<string> col3)
        {
            DataType.DataType dataType;
            using (StreamReader sr = new StreamReader(filePath, Encoding.Default))
            {
                string strDataLine;
                while (!string.IsNullOrEmpty(strDataLine = sr.ReadLine()))
                {
                    string[] strArrDataLine = strDataLine.Split(',');
                    if (strArrDataLine.Length == 5)
                    {
                        col0.Add(strArrDataLine[0]);
                        col1.Add(strArrDataLine[2]);
                        col2.Add(strArrDataLine[3]);
                        col3.Add(strArrDataLine[4]);
                    }
                    else
                    {
                        col0.Add(strArrDataLine[0]);
                        col1.Add(strArrDataLine[1]);
                        col2.Add(strArrDataLine[2]);
                    }
                }
            }
            if (col3.Count >= 1)
            {
                dataType = DataType.DataType.TotalStation;
            }
            else
            {
                try
                {
                    Convert.ToDouble(col1[0]);
                    dataType = DataType.DataType.Gauss;
                }
                catch
                {
                    dataType = DataType.DataType.Geodetic;
                }
            }
            return dataType;
        }
    }
}
