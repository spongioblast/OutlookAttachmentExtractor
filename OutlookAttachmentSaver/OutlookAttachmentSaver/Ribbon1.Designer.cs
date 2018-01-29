namespace OutlookAttachmentSaver
{
    partial class AE : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AE()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.buttonFolderToWatch = this.Factory.CreateRibbonButton();
            this.buttonFolderForAttachments = this.Factory.CreateRibbonButton();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.label3 = this.Factory.CreateRibbonLabel();
            this.label4 = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "AE";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.label1);
            this.group1.Items.Add(this.label3);
            this.group1.Label = "Desc";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.buttonFolderToWatch);
            this.group2.Items.Add(this.buttonFolderForAttachments);
            this.group2.Label = "Settings";
            this.group2.Name = "group2";
            // 
            // buttonFolderToWatch
            // 
            this.buttonFolderToWatch.Label = "Select mail folder to watch";
            this.buttonFolderToWatch.Name = "buttonFolderToWatch";
            this.buttonFolderToWatch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonFolderToWatch_Click);
            // 
            // buttonFolderForAttachments
            // 
            this.buttonFolderForAttachments.Label = "Select folder for extraction";
            this.buttonFolderForAttachments.Name = "buttonFolderForAttachments";
            this.buttonFolderForAttachments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonFolderForAttachments_Click);
            // 
            // label1
            // 
            this.label1.Label = "Watching the following folder:";
            this.label1.Name = "label1";
            // 
            // group3
            // 
            this.group3.Items.Add(this.label2);
            this.group3.Items.Add(this.label4);
            this.group3.Label = "Info";
            this.group3.Name = "group3";
            // 
            // label2
            // 
            this.label2.Label = "Inbox";
            this.label2.Name = "label2";
            // 
            // label3
            // 
            this.label3.Label = "Selected folder for extraction:";
            this.label3.Name = "label3";
            // 
            // label4
            // 
            this.label4.Label = "\"C:\\TempDatenaufnahme\"";
            this.label4.Name = "label4";
            // 
            // AE
            // 
            this.Name = "AE";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonFolderToWatch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonFolderForAttachments;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label4;
    }

    partial class ThisRibbonCollection
    {
        internal AE Ribbon1
        {
            get { return this.GetRibbon<AE>(); }
        }
    }
}
