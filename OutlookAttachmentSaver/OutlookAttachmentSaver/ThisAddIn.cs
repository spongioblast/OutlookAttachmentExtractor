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
        private string folderForAttachments = @"C:\TempDatenaufnahme"; // This should be modifyable by the "select folder for extraction" button in the ribbon
        private string folderToWatch; // Need to define the folder whis is to be monitored -> Should be inbox of Shared Mailbox. This should by modifyable by select mailfolder to watch.

        private Outlook.MAPIFolder inbox;
        private Outlook.Items items;
        private Outlook.NameSpace outlookNameSpace;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            InitializeAddIn();
        }

        private void InboxFolderItemAdded(object Item)
        {
            if (Item is Outlook.MailItem)
            {
                // New mail item in inbox folder
                // MessageBox.Show("you got mail");
                var mailItem = (Outlook.MailItem)Item;
                SaveAttachmentsToDisk(mailItem);
            }
        }

        public void SaveAttachmentsToDisk(Outlook.MailItem message)
        {
            foreach (Outlook.Attachment attachment in message.Attachments)
            {
                string fileName = string.Format("{0}\\{1}", folderForAttachments, attachment.FileName);
                attachment.SaveAsFile(fileName);
            }
        }

        public void InitializeAddIn()
        {
            if (items != null)
            {
                items.ItemAdd -= InboxFolderItemAdded;
            }

            folderForAttachments = string.IsNullOrEmpty(Settings.Default.SaveFolder) ? folderForAttachments : Settings.Default.SaveFolder;
            folderToWatch = Settings.Default.InboxFolder;

            try
            {
                if (!System.IO.Directory.Exists(folderForAttachments))
                {
                    System.IO.Directory.CreateDirectory(folderForAttachments); //create the attachment save as path if it doesnt exist
                }
            }
            catch
            {
                throw new Exception("Cannot create folder: " + folderForAttachments);
            }

            outlookNameSpace = Application.GetNamespace("MAPI");

            if (string.IsNullOrEmpty(folderToWatch))
            {
                inbox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox); //get default inbox

                items = inbox.Items;
                items.ItemAdd += InboxFolderItemAdded; // watch for items getting added
            }
            else
            {
                inbox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent;

                foreach (Outlook.MAPIFolder folder in inbox.Folders)
                {
                    if (folder.Name.Equals(folderToWatch, StringComparison.InvariantCultureIgnoreCase))
                    {
                        items = folder.Items;
                        items.ItemAdd += InboxFolderItemAdded;
                    }
                }
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
