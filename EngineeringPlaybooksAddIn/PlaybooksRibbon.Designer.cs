namespace EngineeringPlaybooksAddIn
{
    partial class PlaybooksRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public PlaybooksRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PlaybooksRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.jciPlaybooksGroup = this.Factory.CreateRibbonGroup();
            this.btnListShapes = this.Factory.CreateRibbonButton();
            this.drawPlaybooksButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.jciPlaybooksGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.jciPlaybooksGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // jciPlaybooksGroup
            // 
            this.jciPlaybooksGroup.Items.Add(this.btnListShapes);
            this.jciPlaybooksGroup.Items.Add(this.drawPlaybooksButton);
            this.jciPlaybooksGroup.Label = "JCI Playbooks";
            this.jciPlaybooksGroup.Name = "jciPlaybooksGroup";
            // 
            // btnListShapes
            // 
            this.btnListShapes.Image = ((System.Drawing.Image)(resources.GetObject("btnListShapes.Image")));
            this.btnListShapes.Label = "Page Properties";
            this.btnListShapes.Name = "btnListShapes";
            this.btnListShapes.ShowImage = true;
            this.btnListShapes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnListShapes_Click);
            // 
            // drawPlaybooksButton
            // 
            this.drawPlaybooksButton.Image = ((System.Drawing.Image)(resources.GetObject("drawPlaybooksButton.Image")));
            this.drawPlaybooksButton.Label = "Draw Playbook";
            this.drawPlaybooksButton.Name = "drawPlaybooksButton";
            this.drawPlaybooksButton.ShowImage = true;
            this.drawPlaybooksButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.drayPlaybooksButton_Click);
            // 
            // PlaybooksRibbon
            // 
            this.Name = "PlaybooksRibbon";
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.PlaybooksRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.jciPlaybooksGroup.ResumeLayout(false);
            this.jciPlaybooksGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup jciPlaybooksGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnListShapes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton drawPlaybooksButton;
    }

    partial class ThisRibbonCollection
    {
        internal PlaybooksRibbon PlaybooksRibbon
        {
            get { return this.GetRibbon<PlaybooksRibbon>(); }
        }
    }
}
