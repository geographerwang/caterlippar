using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using netDxf;
using netDxf.Entities;
using netDxf.Header;

namespace AdjustmentAssistant
{
    class OutputDxf
    {
        internal void OutputTraverse(string p, List<string> col0, List<string> colX, List<string> colY)
        {
            DxfDocument dxfTranvese = new DxfDocument();
            Polyline poly = new Polyline();
            for (int i = 0; i < col0.Count; i++)
            {
                poly.Vertexes.Add(new PolylineVertex(float.Parse(colX[i]), float.Parse(colY[i])));
                Text text = new Text();
                text.Value = col0[i];
                text.BasePoint = new Vector3f(float.Parse(colX[i]), float.Parse(colY[i]), 0f);
                dxfTranvese.AddEntity(text);
            }
            dxfTranvese.AddEntity(poly);
            dxfTranvese.Save(p, DxfVersion.AutoCad2007);//保存为2007格式
        }

        internal void OutputPoint(string p, List<string> col0, List<string> colX, List<string> colY)
        {
            DxfDocument dxfTranvese = new DxfDocument();
            for (int i = 0; i < col0.Count; i++)
            {
                Point point = new Point();
                point.Location = new Vector3f(float.Parse(colX[i]), float.Parse(colY[i]), 0f);
                dxfTranvese.AddEntity(point);
                Text text = new Text();
                text.Value = col0[i];
                text.BasePoint = new Vector3f(float.Parse(colX[i]), float.Parse(colY[i]), 0f);
                dxfTranvese.AddEntity(text);
            }
            dxfTranvese.Save(p, DxfVersion.AutoCad2007);//保存为2007格式
        }
    }
}
