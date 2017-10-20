namespace ExcelCSVImporter
{
    partial class RibbonTools : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonTools()
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
            this.tabCLTools = this.Factory.CreateRibbonTab();
            this.groupMain = this.Factory.CreateRibbonGroup();
            this.butImport = this.Factory.CreateRibbonButton();
            this.butReimport = this.Factory.CreateRibbonButton();
            this.tabCLTools.SuspendLayout();
            this.groupMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabCLTools
            // 
            this.tabCLTools.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabCLTools.Groups.Add(this.groupMain);
            this.tabCLTools.Label = "CL Tools";
            this.tabCLTools.Name = "tabCLTools";
            // 
            // groupMain
            // 
            this.groupMain.Items.Add(this.butImport);
            this.groupMain.Items.Add(this.butReimport);
            this.groupMain.Name = "groupMain";
            // 
            // butImport
            // 
            this.butImport.Label = "Import CSV";
            this.butImport.Name = "butImport";
            this.butImport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butImport_Click);
            // 
            // butReimport
            // 
            this.butReimport.Label = "Reimport Current";
            this.butReimport.Name = "butReimport";
            this.butReimport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butReimport_Click);
            // 
            // RibbonTools
            // 
            this.Name = "RibbonTools";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabCLTools);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonTools_Load);
            this.tabCLTools.ResumeLayout(false);
            this.tabCLTools.PerformLayout();
            this.groupMain.ResumeLayout(false);
            this.groupMain.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabCLTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMain;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butReimport;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonTools RibbonTools
        {
            get { return this.GetRibbon<RibbonTools>(); }
        }
    }
}
