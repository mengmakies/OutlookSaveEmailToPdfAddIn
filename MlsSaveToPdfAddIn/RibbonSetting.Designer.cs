namespace MlsSaveToPdfAddIn
{
    partial class RibbonSetting : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonSetting()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnSaveToPdf = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Pdf From Mln.Com";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnSaveToPdf);
            this.group1.Label = "设置";
            this.group1.Name = "group1";
            // 
            // btnSaveToPdf
            // 
            this.btnSaveToPdf.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSaveToPdf.Image = global::MlsSaveToPdfAddIn.Properties.Resources.pdf_64px_1176741_easyicon_net;
            this.btnSaveToPdf.Label = "Convert email of selected folder";
            this.btnSaveToPdf.Name = "btnSaveToPdf";
            this.btnSaveToPdf.ShowImage = true;
            this.btnSaveToPdf.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveToPdf_Click);
            // 
            // RibbonSetting
            // 
            this.Name = "RibbonSetting";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonSetting_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        public Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        public Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveToPdf;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonSetting RibbonSetting
        {
            get { return this.GetRibbon<RibbonSetting>(); }
        }
    }
}
