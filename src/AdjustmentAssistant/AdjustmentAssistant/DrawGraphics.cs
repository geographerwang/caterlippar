using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

namespace AdjustmentAssistant
{
    class DrawGraphics
    {
        internal System.Drawing.Bitmap DrawPoint(Panel pnlResult, List<string> col0, List<string> col14, List<string> col15)
        {
            pnlResult.Controls.Clear();
            Point[] point = new Point[col0.Count];
            double xMin = Convert.ToDouble(col14[0]), yMin = Convert.ToDouble(col15[0]);
            double xMax = Convert.ToDouble(col14[0]), yMax = Convert.ToDouble(col15[0]);
            for (int i = 1; i < col0.Count; i++)
            {
                if (xMin > Convert.ToDouble(col14[i]))
                {
                    xMin = Convert.ToDouble(col14[i]);
                }
                if (xMax < Convert.ToDouble(col14[i]))
                {
                    xMax = Convert.ToDouble(col14[i]);
                }
                if (yMin > Convert.ToDouble(col15[i]))
                {
                    yMin = Convert.ToDouble(col15[i]);
                }
                if (yMax < Convert.ToDouble(col15[i]))
                {
                    yMax = Convert.ToDouble(col15[i]);
                }
            }
            double width = (xMax - xMin) / 600;
            double heigth = (yMax - yMin) / 400;
            double bi;
            double scale;
            if (width > heigth)
            {
                bi = width;
                scale = 600;
            }
            else
            {
                bi = heigth;
                scale = 400;
            }
            for (int i = 0; i < col0.Count; i++)
            {
                point[i].X = (int)((Convert.ToDouble(col14[i]) - xMin) / bi + 100);
                point[i].Y = 500 - (int)((Convert.ToDouble(col15[i]) - yMin) / bi);
            }
            Bitmap b = new Bitmap(pnlResult.Width, pnlResult.Height);
            pnlResult.BackgroundImage = b;
            Graphics g = Graphics.FromImage(b);
            g.Clear(Color.White);
            Pen pen = new Pen(Color.Red);
            g.DrawString("比例尺=1:" + (bi * scale * 100 / 0.0226917).ToString("#.000"), new Font("宋体", 12), new SolidBrush(Color.Black), 45f, 45f);
            for (int i = 0; i < col0.Count; i++)
            {
                if (i == 0 || i == 1 || i == col0.Count - 2 || i == col0.Count - 1)
                {
                    Point[] p = {
                                    new Point(point[i].X-4,point[i].Y+3),
                                    new Point(point[i].X+4,point[i].Y+3),
                                    new Point(point[i].X,point[i].Y-5)
                                };
                    g.DrawPolygon(pen, p);
                }
                else
                {
                    g.DrawRectangle(pen, point[i].X - 3, point[i].Y - 3, 7, 7);
                }
                g.DrawString(col0[i], new Font("宋体", 10), new SolidBrush(Color.Black), point[i].X + 5, point[i].Y - 15);
            }
            g.Dispose();
            return b;
        }

        internal Bitmap DrawLine(System.Windows.Forms.Panel pnlResult, List<string> col0, List<string> col11, List<string> col12, List<string> col13, List<string> col14, List<string> col15)
        {
            pnlResult.Controls.Clear();
            Point[] point = new Point[col0.Count];
            double xMin = Convert.ToDouble(col14[0]), yMin = Convert.ToDouble(col15[0]);
            double xMax = Convert.ToDouble(col14[0]), yMax = Convert.ToDouble(col15[0]);
            for (int i = 1; i < col0.Count; i++)
            {
                if (xMin > Convert.ToDouble(col14[i]))
                {
                    xMin = Convert.ToDouble(col14[i]);
                }
                if (xMax < Convert.ToDouble(col14[i]))
                {
                    xMax = Convert.ToDouble(col14[i]);
                }
                if (yMin > Convert.ToDouble(col15[i]))
                {
                    yMin = Convert.ToDouble(col15[i]);
                }
                if (yMax < Convert.ToDouble(col15[i]))
                {
                    yMax = Convert.ToDouble(col15[i]);
                }
            }
            double width = (xMax - xMin) / 600;
            double heigth = (yMax - yMin) / 400;
            double bi;
            double scale;
            if (width > heigth)
            {
                bi = width;
                scale = 600;
            }
            else
            {
                bi = heigth;
                scale = 400;
            }
            for (int i = 0; i < col0.Count; i++)
            {
                point[i].X = (int)((Convert.ToDouble(col14[i]) - xMin) / bi + 100);
                point[i].Y = 500 - (int)((Convert.ToDouble(col15[i]) - yMin) / bi);
            }
            Bitmap b = new Bitmap(pnlResult.Width, pnlResult.Height);
            pnlResult.BackgroundImage = b;
            Graphics g = Graphics.FromImage(b);
            g.Clear(Color.White);
            Pen pen = new Pen(Color.Red);
            g.DrawString("比例尺=1:" + (bi * scale * 100 / 0.0226917).ToString("#.000") + " 椭圆比例=1:" + (bi * scale * 100 / 0.0226917 / 500).ToString("#.000"), new Font("宋体", 12), new SolidBrush(Color.Black), 45f, 45f);
            g.DrawLines(new Pen(Color.Black), point);
            g.DrawLine(pen, point[0], point[1]);
            g.DrawLine(pen, point[point.Length - 2], point[point.Length - 1]);
            for (int i = 0; i < col0.Count; i++)
            {
                if (i == 0 || i == 1 || i == col0.Count - 2 || i == col0.Count - 1)
                {
                    Point[] p = {
                                    new Point(point[i].X-4,point[i].Y+3),
                                    new Point(point[i].X+4,point[i].Y+3),
                                    new Point(point[i].X,point[i].Y-5)
                                };
                    g.DrawPolygon(pen, p);
                }
                g.DrawString(col0[i], new Font("宋体", 10), new SolidBrush(Color.Black), point[i].X + 5, point[i].Y - 15);
            }

            for (int i = 2; i < col0.Count - 2; i++)
            {
                int aS = (int)(Convert.ToDouble(col11[i - 2]) * 500);
                int bS = (int)(Convert.ToDouble(col12[i - 2]) * 500);
                g.DrawEllipse(pen, point[i].X - aS, point[i].Y - bS, 2 * aS, 2 * bS);
            }
            g.Dispose();
            return b;
        }
    }
}
