using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Management;

namespace 平差助手注册机
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ManagementClass mc = new ManagementClass("Win32_Processor");
            ManagementObjectCollection moc = mc.GetInstances();
            String strCpuID = null;
            foreach (ManagementObject mo in moc)
            {
                strCpuID = mo.Properties["ProcessorId"].Value.ToString();
                break;
            }
            System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create();
            byte[] md5Buffer = md5.ComputeHash(System.Text.Encoding.Default.GetBytes(strCpuID));
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
            textBox1.Text = sb.ToString();
            textBox1.Focus();
        }
    }
}
