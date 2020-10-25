namespace RemoteEmailManager
{
    partial class RemoteRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RemoteRibbon()
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
            this.RemoteCopy = this.Factory.CreateRibbonGroup();
            this.btnCopy = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.RemoteCopy.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail";
            this.tab1.Groups.Add(this.RemoteCopy);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            // 
            // RemoteCopy
            // 
            this.RemoteCopy.Items.Add(this.btnCopy);
            this.RemoteCopy.Label = "Remote Copy";
            this.RemoteCopy.Name = "RemoteCopy";
            // 
            // btnCopy
            // 
            this.btnCopy.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCopy.Description = "Click on button to copy selected email.";
            this.btnCopy.Label = "Copy";
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.OfficeImageId = "ClickToRunApplyUpdates";
            this.btnCopy.ShowImage = true;
            this.btnCopy.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopy_Click);
            // 
            // RemoteRibbon
            // 
            this.Name = "RemoteRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RemoteRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.RemoteCopy.ResumeLayout(false);
            this.RemoteCopy.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup RemoteCopy;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopy;
    }

    partial class ThisRibbonCollection
    {
        internal RemoteRibbon RemoteRibbon
        {
            get { return this.GetRibbon<RemoteRibbon>(); }
        }
    }
}
