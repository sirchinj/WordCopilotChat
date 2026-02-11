namespace WordCopilotChat
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.checkBoxLog = this.Factory.CreateRibbonCheckBox();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btnModel = this.Factory.CreateRibbonButton();
            this.btnPrompt = this.Factory.CreateRibbonButton();
            this.btnDoc = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "WordCopilotChat";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnModel);
            this.group2.Items.Add(this.separator1);
            this.group2.Items.Add(this.btnPrompt);
            this.group2.Items.Add(this.separator2);
            this.group2.Items.Add(this.btnDoc);
            this.group2.Items.Add(this.separator3);
            this.group2.Items.Add(this.checkBoxLog);
            this.group2.Name = "group2";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // checkBoxLog
            // 
            this.checkBoxLog.Label = "日志存储";
            this.checkBoxLog.Name = "checkBoxLog";
            this.checkBoxLog.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBoxLog_Click);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::WordCopilotChat.Properties.Resources.uil__robot_Blue;
            this.button1.Label = "打开对话页";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // btnModel
            // 
            this.btnModel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnModel.Image = global::WordCopilotChat.Properties.Resources.uil__setting_Blue;
            this.btnModel.Label = "模型管理";
            this.btnModel.Name = "btnModel";
            this.btnModel.ShowImage = true;
            this.btnModel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModel_Click);
            // 
            // btnPrompt
            // 
            this.btnPrompt.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPrompt.Image = global::WordCopilotChat.Properties.Resources.uil__wrench_Blue;
            this.btnPrompt.Label = "提示词管理";
            this.btnPrompt.Name = "btnPrompt";
            this.btnPrompt.ShowImage = true;
            this.btnPrompt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrompt_Click);
            // 
            // btnDoc
            // 
            this.btnDoc.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDoc.Image = global::WordCopilotChat.Properties.Resources.uil__book_reader_Blue;
            this.btnDoc.Label = "文档管理";
            this.btnDoc.Name = "btnDoc";
            this.btnDoc.ShowImage = true;
            this.btnDoc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDoc_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModel;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrompt;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDoc;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxLog;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
