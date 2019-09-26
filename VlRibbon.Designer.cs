namespace Vendorlink.Excel
{
    partial class VlRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public VlRibbon()
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
            this.ribbonTabVendorLink = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.refreshListButton = this.Factory.CreateRibbonButton();
            this.QueryDropwdown = this.Factory.CreateRibbonDropDown();
            this.btnFillSheet = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.LoginButton = this.Factory.CreateRibbonButton();
            this.ribbonTabVendorLink.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // ribbonTabVendorLink
            // 
            this.ribbonTabVendorLink.Groups.Add(this.group2);
            this.ribbonTabVendorLink.Groups.Add(this.group1);
            this.ribbonTabVendorLink.Label = "VendorLink";
            this.ribbonTabVendorLink.Name = "ribbonTabVendorLink";
            // 
            // group1
            // 
            this.group1.Items.Add(this.refreshListButton);
            this.group1.Items.Add(this.QueryDropwdown);
            this.group1.Items.Add(this.btnFillSheet);
            this.group1.Label = "Queries";
            this.group1.Name = "group1";
            // 
            // refreshListButton
            // 
            this.refreshListButton.Label = "Refresh queries";
            this.refreshListButton.Name = "refreshListButton";
            this.refreshListButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RefreshListButton_Click);
            // 
            // QueryDropwdown
            // 
            this.QueryDropwdown.Label = "Query";
            this.QueryDropwdown.Name = "QueryDropwdown";
            // 
            // btnFillSheet
            // 
            this.btnFillSheet.Label = "Fill Sheet";
            this.btnFillSheet.Name = "btnFillSheet";
            this.btnFillSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnFillSheet_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.LoginButton);
            this.group2.Label = "Account";
            this.group2.Name = "group2";
            // 
            // LoginButton
            // 
            this.LoginButton.Label = "Login";
            this.LoginButton.Name = "LoginButton";
            this.LoginButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LoginButton_Click);
            // 
            // VlRibbon
            // 
            this.Name = "VlRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.ribbonTabVendorLink);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.VlRibbon_Load);
            this.ribbonTabVendorLink.ResumeLayout(false);
            this.ribbonTabVendorLink.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ribbonTabVendorLink;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown QueryDropwdown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton refreshListButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFillSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LoginButton;
    }

    partial class ThisRibbonCollection
    {
        internal VlRibbon VlRibbon
        {
            get { return this.GetRibbon<VlRibbon>(); }
        }
    }
}
