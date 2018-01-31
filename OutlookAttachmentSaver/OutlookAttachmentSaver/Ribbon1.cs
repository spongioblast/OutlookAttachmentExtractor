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
            var folder = Globals.ThisAddIn.Application.Session.PickFolder();
            Settings.Default.InboxFolder = folder.Name;
            Settings.Default.Save();
            Globals.ThisAddIn.InitializeAddIn();
        }

        private void buttonFolderForAttachments_Click(object sender, RibbonControlEventArgs e)
        {
            var fbd = new FolderBrowserDialog();

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                Settings.Default.SaveFolder = fbd.SelectedPath;
                Settings.Default.Save();
            }

            Globals.ThisAddIn.InitializeAddIn();
        }

        private void checkBoxEnable_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.Default.AutoExtractStatus = checkBoxEnable.Checked;
            Settings.Default.Save();
            //Globals.ThisAddIn.InitializeAddIn();
        }
    }
}
