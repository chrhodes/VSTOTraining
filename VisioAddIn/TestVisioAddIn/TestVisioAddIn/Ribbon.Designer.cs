namespace TestVisioAddIn
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.rtVisioAddInTemplate = this.Factory.CreateRibbonTab();
            this.rgDocumentActions = this.Factory.CreateRibbonGroup();
            this.rlDocumentActions = this.Factory.CreateRibbonLabel();
            this.rgPageActions = this.Factory.CreateRibbonGroup();
            this.rlPageActions = this.Factory.CreateRibbonLabel();
            this.rgLayerActions = this.Factory.CreateRibbonGroup();
            this.rlLayerActions = this.Factory.CreateRibbonLabel();
            this.rgShapeActions = this.Factory.CreateRibbonGroup();
            this.rlShapeActions = this.Factory.CreateRibbonLabel();
            this.rgDebug = this.Factory.CreateRibbonGroup();
            this.btnDebugWindow = this.Factory.CreateRibbonButton();
            this.btnWatchWindow = this.Factory.CreateRibbonButton();
            this.rcbEnableAppEvents = this.Factory.CreateRibbonCheckBox();
            this.rcbDisplayEvents = this.Factory.CreateRibbonCheckBox();
            this.rcbDisplayChattyEvents = this.Factory.CreateRibbonCheckBox();
            this.rcbDeveloperUIMode = this.Factory.CreateRibbonCheckBox();
            this.grpHelp = this.Factory.CreateRibbonGroup();
            this.btnAddInInfo = this.Factory.CreateRibbonButton();
            this.btnDeveloperMode = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.rtVisioAddInTemplate.SuspendLayout();
            this.rgDocumentActions.SuspendLayout();
            this.rgPageActions.SuspendLayout();
            this.rgLayerActions.SuspendLayout();
            this.rgShapeActions.SuspendLayout();
            this.rgDebug.SuspendLayout();
            this.grpHelp.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // rtVisioAddInTemplate
            // 
            this.rtVisioAddInTemplate.Groups.Add(this.rgDocumentActions);
            this.rtVisioAddInTemplate.Groups.Add(this.rgPageActions);
            this.rtVisioAddInTemplate.Groups.Add(this.rgLayerActions);
            this.rtVisioAddInTemplate.Groups.Add(this.rgShapeActions);
            this.rtVisioAddInTemplate.Groups.Add(this.rgDebug);
            this.rtVisioAddInTemplate.Groups.Add(this.grpHelp);
            this.rtVisioAddInTemplate.Label = "TestVisioAddIn";
            this.rtVisioAddInTemplate.Name = "rtVisioAddInTemplate";
            // 
            // rgDocumentActions
            // 
            this.rgDocumentActions.Items.Add(this.rlDocumentActions);
            this.rgDocumentActions.Label = "Document Actions";
            this.rgDocumentActions.Name = "rgDocumentActions";
            // 
            // rlDocumentActions
            // 
            this.rlDocumentActions.Label = "Document Actions Go Here";
            this.rlDocumentActions.Name = "rlDocumentActions";
            // 
            // rgPageActions
            // 
            this.rgPageActions.Items.Add(this.rlPageActions);
            this.rgPageActions.Label = "Page Actions";
            this.rgPageActions.Name = "rgPageActions";
            // 
            // rlPageActions
            // 
            this.rlPageActions.Label = "Page Actions Go Here";
            this.rlPageActions.Name = "rlPageActions";
            // 
            // rgLayerActions
            // 
            this.rgLayerActions.Items.Add(this.rlLayerActions);
            this.rgLayerActions.Label = "Layer Actions";
            this.rgLayerActions.Name = "rgLayerActions";
            // 
            // rlLayerActions
            // 
            this.rlLayerActions.Label = " Layer Actions Go Here";
            this.rlLayerActions.Name = "rlLayerActions";
            // 
            // rgShapeActions
            // 
            this.rgShapeActions.Items.Add(this.rlShapeActions);
            this.rgShapeActions.Label = "Shape Actions";
            this.rgShapeActions.Name = "rgShapeActions";
            // 
            // rlShapeActions
            // 
            this.rlShapeActions.Label = "Shape Actions Go Here";
            this.rlShapeActions.Name = "rlShapeActions";
            // 
            // rgDebug
            // 
            this.rgDebug.Items.Add(this.btnDebugWindow);
            this.rgDebug.Items.Add(this.btnWatchWindow);
            this.rgDebug.Items.Add(this.rcbEnableAppEvents);
            this.rgDebug.Items.Add(this.rcbDisplayEvents);
            this.rgDebug.Items.Add(this.rcbDisplayChattyEvents);
            this.rgDebug.Items.Add(this.rcbDeveloperUIMode);
            this.rgDebug.Label = "Debug";
            this.rgDebug.Name = "rgDebug";
            this.rgDebug.Visible = false;
            // 
            // btnDebugWindow
            // 
            this.btnDebugWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDebugWindow.Image = global::TestVisioAddIn.Properties.Resources.DebugWindow;
            this.btnDebugWindow.Label = "Debug Window";
            this.btnDebugWindow.Name = "btnDebugWindow";
            this.btnDebugWindow.ShowImage = true;
            this.btnDebugWindow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDebugWindow_Click);
            // 
            // btnWatchWindow
            // 
            this.btnWatchWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnWatchWindow.Image = global::TestVisioAddIn.Properties.Resources.WatchWindow;
            this.btnWatchWindow.Label = "Watch Window";
            this.btnWatchWindow.Name = "btnWatchWindow";
            this.btnWatchWindow.ShowImage = true;
            this.btnWatchWindow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWatchWindow_Click);
            // 
            // rcbEnableAppEvents
            // 
            this.rcbEnableAppEvents.Label = "Enable App Events";
            this.rcbEnableAppEvents.Name = "rcbEnableAppEvents";
            this.rcbEnableAppEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rcbEnableAppEvents_Click);
            // 
            // rcbDisplayEvents
            // 
            this.rcbDisplayEvents.Label = "Display Events";
            this.rcbDisplayEvents.Name = "rcbDisplayEvents";
            this.rcbDisplayEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rcbDisplayEvents_Click);
            // 
            // rcbDisplayChattyEvents
            // 
            this.rcbDisplayChattyEvents.Label = "Display Chatty Events";
            this.rcbDisplayChattyEvents.Name = "rcbDisplayChattyEvents";
            this.rcbDisplayChattyEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rcbDisplayChattyEvents_Click);
            // 
            // rcbDeveloperUIMode
            // 
            this.rcbDeveloperUIMode.Label = "DeveloperUIMode";
            this.rcbDeveloperUIMode.Name = "rcbDeveloperUIMode";
            this.rcbDeveloperUIMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToggleDeveloperUIMode_Click);
            // 
            // grpHelp
            // 
            this.grpHelp.Items.Add(this.btnAddInInfo);
            this.grpHelp.Items.Add(this.btnDeveloperMode);
            this.grpHelp.Label = "Help";
            this.grpHelp.Name = "grpHelp";
            // 
            // btnAddInInfo
            // 
            this.btnAddInInfo.Label = "AddIn Info";
            this.btnAddInInfo.Name = "btnAddInInfo";
            this.btnAddInInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDisplayAddInInfo_Click);
            // 
            // btnDeveloperMode
            // 
            this.btnDeveloperMode.Label = "Developer Mode";
            this.btnDeveloperMode.Name = "btnDeveloperMode";
            this.btnDeveloperMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToggleDeveloperMode_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.rtVisioAddInTemplate);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.rtVisioAddInTemplate.ResumeLayout(false);
            this.rtVisioAddInTemplate.PerformLayout();
            this.rgDocumentActions.ResumeLayout(false);
            this.rgDocumentActions.PerformLayout();
            this.rgPageActions.ResumeLayout(false);
            this.rgPageActions.PerformLayout();
            this.rgLayerActions.ResumeLayout(false);
            this.rgLayerActions.PerformLayout();
            this.rgShapeActions.ResumeLayout(false);
            this.rgShapeActions.PerformLayout();
            this.rgDebug.ResumeLayout(false);
            this.rgDebug.PerformLayout();
            this.grpHelp.ResumeLayout(false);
            this.grpHelp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab rtVisioAddInTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgDocumentActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgPageActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel rlDocumentActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel rlPageActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgLayerActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel rlLayerActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgShapeActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel rlShapeActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgDebug;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDebugWindow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWatchWindow;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbEnableAppEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbDisplayEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbDisplayChattyEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbDeveloperUIMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddInInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeveloperMode;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
