using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AdjustmentAssistant
{
    public partial class FormAbout : Form
    {
        public FormAbout()
        {
            InitializeComponent();
        }

        public FormAbout(string isRegedit)
            : this()
        {
            this.lblRegedit.Text = isRegedit;
        }

        private void btnAboutSure_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
