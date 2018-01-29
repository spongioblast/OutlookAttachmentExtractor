using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookAttachmentSaver
{
    public partial class ThisAddIn
    {
        string attachmentSaveLocation = "C:\\TempDatenaufnahme"; // This should be modifyable by the "select folder for extraction" button in the ribbon
        // Need to define the folder whis is to be monitored -> Should be inbox of Shared Mailbox. This should by modifyable by select mailfolder to watch.

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.IO.Directory.CreateDirectory(attachmentSaveLocation); //create the attachment save as path if it doesnt exist
            Outlook.MAPIFolder inbox = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox); //get default inbox
            inbox.Items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(InboxFolderItemAdded); // watch for items getting added
        }

        private void InboxFolderItemAdded(object Item)
        {
            if (Item is Outlook.MailItem)
            {
                // New mail item in inbox folder
                // MessageBox.Show("you got mail");
                Outlook.MailItem mailItem = (Outlook.MailItem)Item;
                SaveAttachmentsToDisk(mailItem);
            }
        }

        public void SaveAttachmentsToDisk(Outlook.MailItem message)
        {
            foreach (Outlook.Attachment attachment in message.Attachments)
            {
                string fileName = string.Format("{0}\\{1}", attachmentSaveLocation, attachment.FileName);
                attachment.SaveAsFile(fileName);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
