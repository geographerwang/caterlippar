using System;
using System.Windows.Forms;

namespace AdjustmentAssistant
{
    public static class ClearPanel
    {
        public static void Clear(Panel pnl)
        {
            pnl.Controls.Clear();
        }
    }
}
