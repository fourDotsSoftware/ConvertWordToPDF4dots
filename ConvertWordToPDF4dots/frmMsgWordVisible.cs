using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ConvertWordToPDF4dots
{
    public partial class frmMsgWordVisible : ConvertWordToPDF4dots.CustomForm
    {
        public frmMsgWordVisible()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.MsgWordVisible = !chkDoNotShowAgain.Checked;

            this.DialogResult = DialogResult.OK;
        }
    }
}
