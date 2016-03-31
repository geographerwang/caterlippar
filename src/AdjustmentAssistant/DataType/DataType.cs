using System;

namespace DataType
{
    /// <summary>
    /// 读取的数据类型，分别对应近似平差，全站仪文件，高斯坐标，经纬度，手动录入数据
    /// </summary>
    public enum DataType
    {
        ApproximateAdjustment,
        TotalStation,
        Gauss,
        Geodetic,
        HandMade
    }

    /// <summary>
    /// 近似平差的数据类型，分别对应闭附合导线，支导线，单面水准，双面水准
    /// </summary>
    public enum Data
    {
        ConnectingTraverse,
        OpenTraverse,
        SingleRule,
        DoubleRule
    }

    /// <summary>
    /// 观测角方向，分别对应左角和右角
    /// </summary>
    public enum LeftOrRight
    {
        Left,
        Right
    }

    /// <summary>
    /// 水准观测权重，分别对应测站数，距离
    /// </summary>
    public enum LevelingWeight
    {
        PointCount,
        Lenght
    }
}
