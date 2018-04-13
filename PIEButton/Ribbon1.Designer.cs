namespace PIEButton
{
    partial class PIERibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public PIERibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PIERibbon));
            this.pieTab1 = this.Factory.CreateRibbonTab();
            this.pieGroup1 = this.Factory.CreateRibbonGroup();
            this.pieButton1 = this.Factory.CreateRibbonButton();
            this.pieTab2 = this.Factory.CreateRibbonTab();
            this.pieGroup2 = this.Factory.CreateRibbonGroup();
            this.pieButton2 = this.Factory.CreateRibbonButton();
            this.pieTab1.SuspendLayout();
            this.pieGroup1.SuspendLayout();
            this.pieTab2.SuspendLayout();
            this.pieGroup2.SuspendLayout();
            this.SuspendLayout();
            // 
            // pieTab1
            // 
            this.pieTab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.pieTab1.ControlId.OfficeId = "TabMail";
            this.pieTab1.Groups.Add(this.pieGroup1);
            this.pieTab1.Label = "TabMail";
            this.pieTab1.Name = "pieTab1";
            // 
            // pieGroup1
            // 
            this.pieGroup1.Items.Add(this.pieButton1);
            this.pieGroup1.Label = "LogRhythm";
            this.pieGroup1.Name = "pieGroup1";
            // 
            // pieButton1
            // 
            this.pieButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.pieButton1.Image = ((System.Drawing.Image)(resources.GetObject("pieButton1.Image")));
            this.pieButton1.Label = "Report Phishing";
            this.pieButton1.Name = "pieButton1";
            this.pieButton1.ScreenTip = "Report Phishing E-Mail";
            this.pieButton1.ShowImage = true;
            this.pieButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pieButton1_Click);
            // 
            // pieTab2
            // 
            this.pieTab2.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.pieTab2.ControlId.OfficeId = "TabReadMessage";
            this.pieTab2.Groups.Add(this.pieGroup2);
            this.pieTab2.Label = "TabReadMessage";
            this.pieTab2.Name = "pieTab2";
            // 
            // pieGroup2
            // 
            this.pieGroup2.Items.Add(this.pieButton2);
            this.pieGroup2.Label = "LogRhythm";
            this.pieGroup2.Name = "pieGroup2";
            // 
            // pieButton2
            // 
            this.pieButton2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.pieButton2.Image = ((System.Drawing.Image)(resources.GetObject("pieButton2.Image")));
            this.pieButton2.Label = "Report Phishing";
            this.pieButton2.Name = "pieButton2";
            this.pieButton2.ScreenTip = "Report Phishing E-Mail";
            this.pieButton2.ShowImage = true;
            this.pieButton2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pieButton2_Click);
            // 
            // PIERibbon
            // 
            this.Name = "PIERibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.pieTab1);
            this.Tabs.Add(this.pieTab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.pieTab1.ResumeLayout(false);
            this.pieTab1.PerformLayout();
            this.pieGroup1.ResumeLayout(false);
            this.pieGroup1.PerformLayout();
            this.pieTab2.ResumeLayout(false);
            this.pieTab2.PerformLayout();
            this.pieGroup2.ResumeLayout(false);
            this.pieGroup2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup pieGroup1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab pieTab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup pieGroup2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pieButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pieButton2;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab pieTab1;
    }

    partial class ThisRibbonCollection
    {
        internal PIERibbon Ribbon1
        {
            get { return this.GetRibbon<PIERibbon>(); }
        }
    }
}
