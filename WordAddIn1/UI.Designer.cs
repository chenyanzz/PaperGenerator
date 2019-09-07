namespace WordAddIn1
{
    partial class UI : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public UI()
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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.aa = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn_SelectFormatDocx = this.Factory.CreateRibbonButton();
            this.btn_SelectMdFile = this.Factory.CreateRibbonButton();
            this.btn_BuildDocx = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.cb_DocxSelected = this.Factory.CreateRibbonCheckBox();
            this.cb_MdSelected = this.Factory.CreateRibbonCheckBox();
            this.openFileDialog_docx = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialog_md = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.aa.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // aa
            // 
            this.aa.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.aa.Groups.Add(this.group1);
            this.aa.Groups.Add(this.group2);
            this.aa.Label = "Paper Generator";
            this.aa.Name = "aa";
            this.aa.Tag = "";
            // 
            // group1
            // 
            this.group1.DialogLauncher = ribbonDialogLauncherImpl1;
            this.group1.Items.Add(this.btn_SelectFormatDocx);
            this.group1.Items.Add(this.btn_SelectMdFile);
            this.group1.Items.Add(this.btn_BuildDocx);
            this.group1.Label = "Paper Generate";
            this.group1.Name = "group1";
            // 
            // btn_SelectFormatDocx
            // 
            this.btn_SelectFormatDocx.Label = "SelectFormatDocx";
            this.btn_SelectFormatDocx.Name = "btn_SelectFormatDocx";
            this.btn_SelectFormatDocx.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SelectFormatDocx_Click);
            // 
            // btn_SelectMdFile
            // 
            this.btn_SelectMdFile.Label = "SelectMdFile";
            this.btn_SelectMdFile.Name = "btn_SelectMdFile";
            this.btn_SelectMdFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SelectMdFile_Click);
            // 
            // btn_BuildDocx
            // 
            this.btn_BuildDocx.Label = "BuildDocx";
            this.btn_BuildDocx.Name = "btn_BuildDocx";
            this.btn_BuildDocx.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_BuildDocx_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.cb_DocxSelected);
            this.group2.Items.Add(this.cb_MdSelected);
            this.group2.Label = "status";
            this.group2.Name = "group2";
            // 
            // cb_DocxSelected
            // 
            this.cb_DocxSelected.Label = "DocxSelected";
            this.cb_DocxSelected.Name = "cb_DocxSelected";
            // 
            // cb_MdSelected
            // 
            this.cb_MdSelected.Label = "MdSelected";
            this.cb_MdSelected.Name = "cb_MdSelected";
            // 
            // openFileDialog_docx
            // 
            this.openFileDialog_docx.Filter = "Word Xml|*.docx";
            this.openFileDialog_docx.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog_docx_FileOk);
            // 
            // openFileDialog_md
            // 
            this.openFileDialog_md.Filter = "Markdown Files|*.md";
            this.openFileDialog_md.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog_md_FileOk);
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.Filter = "Word Xml|*.docx";
            this.saveFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog_FileOk);
            // 
            // UI
            // 
            this.Name = "UI";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.aa);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.UI_Load);
            this.aa.ResumeLayout(false);
            this.aa.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab aa;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SelectFormatDocx;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SelectMdFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_BuildDocx;
        private System.Windows.Forms.OpenFileDialog openFileDialog_docx;
        private System.Windows.Forms.OpenFileDialog openFileDialog_md;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cb_DocxSelected;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cb_MdSelected;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
    }

    partial class ThisRibbonCollection
    {
        internal UI UI
        {
            get { return this.GetRibbon<UI>(); }
        }
    }
}
