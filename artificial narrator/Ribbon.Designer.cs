
namespace artificial_narrator
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupArtificalNarrator = this.Factory.CreateRibbonGroup();
            this.TestSpeech = this.Factory.CreateRibbonButton();
            this.InsertNarration = this.Factory.CreateRibbonButton();
            this.InsertNarrationAll = this.Factory.CreateRibbonButton();
            this.VoiceListBox = this.Factory.CreateRibbonDropDown();
            this.RateBox = this.Factory.CreateRibbonDropDown();
            this.tab1.SuspendLayout();
            this.groupArtificalNarrator.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupArtificalNarrator);
            this.tab1.Label = "人工ナレーター";
            this.tab1.Name = "tab1";
            // 
            // groupArtificalNarrator
            // 
            this.groupArtificalNarrator.Items.Add(this.TestSpeech);
            this.groupArtificalNarrator.Items.Add(this.InsertNarration);
            this.groupArtificalNarrator.Items.Add(this.InsertNarrationAll);
            this.groupArtificalNarrator.Items.Add(this.VoiceListBox);
            this.groupArtificalNarrator.Items.Add(this.RateBox);
            this.groupArtificalNarrator.Label = "人工ナレーター";
            this.groupArtificalNarrator.Name = "groupArtificalNarrator";
            // 
            // TestSpeech
            // 
            this.TestSpeech.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.TestSpeech.Image = global::artificial_narrator.Properties.Resources.PlaybackPreview_16x;
            this.TestSpeech.Label = "テスト再生";
            this.TestSpeech.Name = "TestSpeech";
            this.TestSpeech.ShowImage = true;
            this.TestSpeech.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestSpeech_Click);
            // 
            // InsertNarration
            // 
            this.InsertNarration.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.InsertNarration.Image = global::artificial_narrator.Properties.Resources.speechballoon_80403;
            this.InsertNarration.Label = "このスライドにナレーションを挿入";
            this.InsertNarration.Name = "InsertNarration";
            this.InsertNarration.ShowImage = true;
            this.InsertNarration.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InsertNarration_Click);
            // 
            // InsertNarrationAll
            // 
            this.InsertNarrationAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.InsertNarrationAll.Image = global::artificial_narrator.Properties.Resources.speechballoons_80420;
            this.InsertNarrationAll.Label = "すべてのスライドにナレーションを挿入";
            this.InsertNarrationAll.Name = "InsertNarrationAll";
            this.InsertNarrationAll.ShowImage = true;
            this.InsertNarrationAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InsertNarration_Click);
            // 
            // VoiceListBox
            // 
            this.VoiceListBox.Label = "音声";
            this.VoiceListBox.Name = "VoiceListBox";
            this.VoiceListBox.SizeString = "Mixrosoft WWWW Desktop";
            // 
            // RateBox
            // 
            this.RateBox.Label = "速度";
            this.RateBox.Name = "RateBox";
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupArtificalNarrator.ResumeLayout(false);
            this.groupArtificalNarrator.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupArtificalNarrator;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton InsertNarration;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TestSpeech;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown VoiceListBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown RateBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton InsertNarrationAll;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
