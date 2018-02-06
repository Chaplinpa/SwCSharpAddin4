using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using SolidWorksTools.File;
using SwCSharpAddin4;

namespace SwCSharpAddin4
{
    public partial class frmDrawConfigSwitch : Form
    {
        public frmDrawConfigSwitch(string[] configNames, string valText)
        {
            InitializeComponent();
            for (int i = 0; i< configNames.Length; i++)
            {
                cmbConfig.Items.Add(configNames[i].ToString());
            }
            cmbConfig.Text = valText;
        }

        public void cmdOK_Click(object sender, EventArgs e) 
        {
            SwAddin.drwChangeConfig(cmbConfig.Text);
            this.Close();
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
