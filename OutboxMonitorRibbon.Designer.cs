
namespace Outbox_Monitor
{
    partial class OutboxMonitorRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public OutboxMonitorRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OutboxMonitorRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Execute = this.Factory.CreateRibbonGroup();
            this.btnProcessOutboxItems = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Execute.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Execute);
            this.tab1.Label = "Outbox Monitor";
            this.tab1.Name = "tab1";
            // 
            // Execute
            // 
            this.Execute.Items.Add(this.btnProcessOutboxItems);
            this.Execute.Label = "Execute";
            this.Execute.Name = "Execute";
            // 
            // btnProcessOutboxItems
            // 
            this.btnProcessOutboxItems.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnProcessOutboxItems.Image = ((System.Drawing.Image)(resources.GetObject("btnProcessOutboxItems.Image")));
            this.btnProcessOutboxItems.Label = "Process Outbox Items";
            this.btnProcessOutboxItems.Name = "btnProcessOutboxItems";
            this.btnProcessOutboxItems.ShowImage = true;
            this.btnProcessOutboxItems.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProcessOutboxItems_Click);
            // 
            // OutboxMonitorRibbon
            // 
            this.Name = "OutboxMonitorRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.OutboxMonitorRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Execute.ResumeLayout(false);
            this.Execute.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Execute;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnProcessOutboxItems;
    }

    partial class ThisRibbonCollection
    {
        internal OutboxMonitorRibbon OutboxMonitorRibbon
        {
            get { return this.GetRibbon<OutboxMonitorRibbon>(); }
        }
    }
}
