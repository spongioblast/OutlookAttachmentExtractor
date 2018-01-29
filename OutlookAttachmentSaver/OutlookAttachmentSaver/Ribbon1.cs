using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace OutlookAttachmentSaver
{
    public partial class AE
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonFolderToWatch_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Sorry, this function  is not yet working. Adrian.");
        }

        private void buttonFolderForAttachments_Click(object sender, RibbonControlEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                MessageBox.Show("Sorry, this function  is not yet working. Adrian.");
        }
    }
}
