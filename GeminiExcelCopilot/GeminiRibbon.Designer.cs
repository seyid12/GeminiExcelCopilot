namespace GeminiExcelCopilot
{
    partial class GeminiRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public GeminiRibbon()
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
            this.tabGemini = this.Factory.CreateRibbonTab();
            this.groupControls = this.Factory.CreateRibbonGroup();
            this.toggleButtonShowPane = this.Factory.CreateRibbonToggleButton();
            this.tabGemini.SuspendLayout();
            this.groupControls.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabGemini
            // 
            this.tabGemini.Groups.Add(this.groupControls);
            this.tabGemini.Label = "Veri Asistanı";
            this.tabGemini.Name = "tabGemini";
            // 
            // groupControls
            // 
            this.groupControls.Items.Add(this.toggleButtonShowPane);
            this.groupControls.Label = "Kontroller";
            this.groupControls.Name = "groupControls";
            // 
            // toggleButtonShowPane
            // 
            this.toggleButtonShowPane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButtonShowPane.Label = "Gemini Veri Asistanını Göster";
            this.toggleButtonShowPane.Name = "toggleButtonShowPane";
            this.toggleButtonShowPane.OfficeImageId = "FindDialog";
            this.toggleButtonShowPane.ShowImage = true;
            this.toggleButtonShowPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonShowPane_Click);
            // 
            // GeminiRibbon
            // 
            this.Name = "GeminiRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabGemini);
            this.tabGemini.ResumeLayout(false);
            this.tabGemini.PerformLayout();
            this.groupControls.ResumeLayout(false);
            this.groupControls.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabGemini;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupControls;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonShowPane;
    }

    partial class ThisRibbonCollection
    {
        internal GeminiRibbon GeminiRibbon
        {
            get { return this.GetRibbon<GeminiRibbon>(); }
        }
    }
}