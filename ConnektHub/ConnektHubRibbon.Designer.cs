namespace Prospecta.ConnektHub
{
    partial class ConnektHubRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ConnektHubRibbon()
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
            this.tabMdoMetadata = this.Factory.CreateRibbonTab();
            this.grpUserDetails = this.Factory.CreateRibbonGroup();
            this.grpModules = this.Factory.CreateRibbonGroup();
            this.grpOptions = this.Factory.CreateRibbonGroup();
            this.grpAction = this.Factory.CreateRibbonGroup();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.btnLogin = this.Factory.CreateRibbonButton();
            this.btnUserDetails = this.Factory.CreateRibbonButton();
            this.btnModules = this.Factory.CreateRibbonButton();
            this.dynamicMenuModules = this.Factory.CreateRibbonMenu();
            this.dynamicMenuAdd = this.Factory.CreateRibbonMenu();
            this.btnFields = this.Factory.CreateRibbonButton();
            this.btnFieldDescriptions = this.Factory.CreateRibbonButton();
            this.btnDropdowns = this.Factory.CreateRibbonButton();
            this.btnFieldNDropdowns = this.Factory.CreateRibbonButton();
            this.btnImport = this.Factory.CreateRibbonButton();
            this.btnExport = this.Factory.CreateRibbonButton();
            this.tabMdoMetadata.SuspendLayout();
            this.grpUserDetails.SuspendLayout();
            this.grpModules.SuspendLayout();
            this.grpOptions.SuspendLayout();
            this.grpAction.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMdoMetadata
            // 
            this.tabMdoMetadata.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabMdoMetadata.Groups.Add(this.grpUserDetails);
            this.tabMdoMetadata.Groups.Add(this.grpModules);
            this.tabMdoMetadata.Groups.Add(this.grpOptions);
            this.tabMdoMetadata.Groups.Add(this.grpAction);
            this.tabMdoMetadata.Label = "MDO Metadata";
            this.tabMdoMetadata.Name = "tabMdoMetadata";
            // 
            // grpUserDetails
            // 
            this.grpUserDetails.Items.Add(this.btnLogin);
            this.grpUserDetails.Items.Add(this.btnUserDetails);
            this.grpUserDetails.Label = "User Details";
            this.grpUserDetails.Name = "grpUserDetails";
            // 
            // grpModules
            // 
            this.grpModules.Items.Add(this.btnModules);
            this.grpModules.Items.Add(this.dynamicMenuModules);
            this.grpModules.Label = "Modules";
            this.grpModules.Name = "grpModules";
            // 
            // grpOptions
            // 
            this.grpOptions.Items.Add(this.dynamicMenuAdd);
            this.grpOptions.Label = "Options";
            this.grpOptions.Name = "grpOptions";
            // 
            // grpAction
            // 
            this.grpAction.Items.Add(this.btnImport);
            this.grpAction.Items.Add(this.btnExport);
            this.grpAction.Label = "Action";
            this.grpAction.Name = "grpAction";
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BackgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.BackgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BackgroundWorker1_RunWorkerCompleted);
            // 
            // btnLogin
            // 
            this.btnLogin.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLogin.Label = "Login";
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.OfficeImageId = "Lock";
            this.btnLogin.ShowImage = true;
            this.btnLogin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnLogin_Click);
            // 
            // btnUserDetails
            // 
            this.btnUserDetails.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUserDetails.Label = "User Image";
            this.btnUserDetails.Name = "btnUserDetails";
            this.btnUserDetails.OfficeImageId = "ContactPictureMenu";
            this.btnUserDetails.ShowImage = true;
            // 
            // btnModules
            // 
            this.btnModules.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnModules.Label = "Modules";
            this.btnModules.Name = "btnModules";
            this.btnModules.OfficeImageId = "ModuleInsert";
            this.btnModules.ShowImage = true;
            this.btnModules.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnModules_Click);
            // 
            // dynamicMenuModules
            // 
            this.dynamicMenuModules.Dynamic = true;
            this.dynamicMenuModules.Label = "Select Module";
            this.dynamicMenuModules.Name = "dynamicMenuModules";
            this.dynamicMenuModules.OfficeImageId = "CreateClassModule";
            this.dynamicMenuModules.ShowImage = true;
            // 
            // dynamicMenuAdd
            // 
            this.dynamicMenuAdd.Dynamic = true;
            this.dynamicMenuAdd.Items.Add(this.btnFields);
            this.dynamicMenuAdd.Items.Add(this.btnFieldDescriptions);
            this.dynamicMenuAdd.Items.Add(this.btnDropdowns);
            this.dynamicMenuAdd.Items.Add(this.btnFieldNDropdowns);
            this.dynamicMenuAdd.Label = "Add";
            this.dynamicMenuAdd.Name = "dynamicMenuAdd";
            this.dynamicMenuAdd.OfficeImageId = "SourceControlAddObjects";
            this.dynamicMenuAdd.ShowImage = true;
            // 
            // btnFields
            // 
            this.btnFields.Label = "Fields";
            this.btnFields.Name = "btnFields";
            this.btnFields.ShowImage = true;
            // 
            // btnFieldDescriptions
            // 
            this.btnFieldDescriptions.Label = "Descriptions";
            this.btnFieldDescriptions.Name = "btnFieldDescriptions";
            this.btnFieldDescriptions.ShowImage = true;
            // 
            // btnDropdowns
            // 
            this.btnDropdowns.Label = "Dropdowns";
            this.btnDropdowns.Name = "btnDropdowns";
            this.btnDropdowns.ShowImage = true;
            // 
            // btnFieldNDropdowns
            // 
            this.btnFieldNDropdowns.Label = "Fields And Dropdowns";
            this.btnFieldNDropdowns.Name = "btnFieldNDropdowns";
            this.btnFieldNDropdowns.ShowImage = true;
            // 
            // btnImport
            // 
            this.btnImport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImport.Label = "Import Fields";
            this.btnImport.Name = "btnImport";
            this.btnImport.OfficeImageId = "TableExportTableToSharePointList";
            this.btnImport.ShowImage = true;
            this.btnImport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnImport_Click);
            // 
            // btnExport
            // 
            this.btnExport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExport.Label = "Export";
            this.btnExport.Name = "btnExport";
            this.btnExport.OfficeImageId = "SaveAndClose";
            this.btnExport.ShowImage = true;
            this.btnExport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnExport_Click);
            // 
            // ConnektHubRibbon
            // 
            this.Name = "ConnektHubRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabMdoMetadata);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ConnektHubRibbon_Load);
            this.tabMdoMetadata.ResumeLayout(false);
            this.tabMdoMetadata.PerformLayout();
            this.grpUserDetails.ResumeLayout(false);
            this.grpUserDetails.PerformLayout();
            this.grpModules.ResumeLayout(false);
            this.grpModules.PerformLayout();
            this.grpOptions.ResumeLayout(false);
            this.grpOptions.PerformLayout();
            this.grpAction.ResumeLayout(false);
            this.grpAction.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabMdoMetadata;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpUserDetails;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpModules;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLogin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUserDetails;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModules;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu dynamicMenuModules;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu dynamicMenuAdd;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpAction;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFields;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFieldDescriptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDropdowns;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFieldNDropdowns;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
    }

    partial class ThisRibbonCollection
    {
        internal ConnektHubRibbon ConnektHubRibbon
        {
            get { return this.GetRibbon<ConnektHubRibbon>(); }
        }
    }
}
