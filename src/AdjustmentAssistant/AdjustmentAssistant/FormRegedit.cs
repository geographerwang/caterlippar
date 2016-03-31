using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Reflection;

namespace AdjustmentAssistant
{
    public partial class FormRegedit : Form
    {
        private string cdKey;

        public FormRegedit()
        {
            InitializeComponent();
        }

        public FormRegedit(string str)
            : this()
        {
            cdKey = str;
        }

        private void WriteKey()
        {
            string keyPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Set.ini");
            using (StreamWriter sw = new StreamWriter(keyPath, false, Encoding.Default))
            {
                sw.WriteLine(this.txtRegedit.Text.Trim());
                if (this.txtRegedit.Text.Trim() == cdKey)
                {
                    MessageBox.Show("注册成功，重启软件后生效");
                }
                else
                {
                    MessageBox.Show("无效的激活码");
                }
            }
        }

        private void btnRegeditSure_Click(object sender, EventArgs e)
        {
            WriteKey();
            this.Close();
        }

        private void btnRegeditCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
