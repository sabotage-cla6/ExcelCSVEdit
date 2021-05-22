namespace CSVDEditAddin
{
  /// <summary>CSVを編集するためのEXCELAddinのリボン</summary>
  partial class CSVEdirRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
  {
    /// <summary>
    /// 必要なデザイナー変数です。
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>コンストラクタ</summary>
    public CSVEdirRibbon()
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
      Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
      Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
      this.csvEditRab = this.Factory.CreateRibbonTab();
      this.CSVFileGroup = this.Factory.CreateRibbonGroup();
      this.openFolderButton = this.Factory.CreateRibbonButton();
      this.saveFolderButton = this.Factory.CreateRibbonButton();
      this.saveFolderAsButton = this.Factory.CreateRibbonButton();
      this.group1 = this.Factory.CreateRibbonGroup();
      this.spearatorEditBox = this.Factory.CreateRibbonEditBox();
      this.extEditBox = this.Factory.CreateRibbonEditBox();
      this.encodingComboBox = this.Factory.CreateRibbonComboBox();
      this.csvEditRab.SuspendLayout();
      this.CSVFileGroup.SuspendLayout();
      this.group1.SuspendLayout();
      this.SuspendLayout();
      // 
      // csvEditRab
      // 
      this.csvEditRab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
      this.csvEditRab.Groups.Add(this.CSVFileGroup);
      this.csvEditRab.Groups.Add(this.group1);
      this.csvEditRab.Label = "CSV";
      this.csvEditRab.Name = "csvEditRab";
      // 
      // CSVFileGroup
      // 
      this.CSVFileGroup.Items.Add(this.openFolderButton);
      this.CSVFileGroup.Items.Add(this.saveFolderButton);
      this.CSVFileGroup.Items.Add(this.saveFolderAsButton);
      this.CSVFileGroup.Label = "ファイル";
      this.CSVFileGroup.Name = "CSVFileGroup";
      // 
      // openFolderButton
      // 
      this.openFolderButton.Label = "フォルダを開く";
      this.openFolderButton.Name = "openFolderButton";
      this.openFolderButton.OfficeImageId = "FileOpen";
      this.openFolderButton.ShowImage = true;
      this.openFolderButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.openFolderButton_Click);
      // 
      // saveFolderButton
      // 
      this.saveFolderButton.Label = "フォルダに上書き保存";
      this.saveFolderButton.Name = "saveFolderButton";
      this.saveFolderButton.OfficeImageId = "FileSave";
      this.saveFolderButton.ShowImage = true;
      this.saveFolderButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveFolderButton_Click);
      // 
      // saveFolderAsButton
      // 
      this.saveFolderAsButton.Label = "フォルダ名を指定して保存";
      this.saveFolderAsButton.Name = "saveFolderAsButton";
      this.saveFolderAsButton.OfficeImageId = "FileSaveAs";
      this.saveFolderAsButton.ShowImage = true;
      this.saveFolderAsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveFolderAsButton_Click);
      // 
      // group1
      // 
      this.group1.Items.Add(this.spearatorEditBox);
      this.group1.Items.Add(this.extEditBox);
      this.group1.Items.Add(this.encodingComboBox);
      this.group1.Label = "ファイル情報";
      this.group1.Name = "group1";
      // 
      // spearatorEditBox
      // 
      this.spearatorEditBox.Label = "区切り文字";
      this.spearatorEditBox.Name = "spearatorEditBox";
      this.spearatorEditBox.Text = ",";
      // 
      // extEditBox
      // 
      this.extEditBox.Label = "ファイル拡張子";
      this.extEditBox.Name = "extEditBox";
      this.extEditBox.Text = "*.csv";
      // 
      // encodingComboBox
      // 
      ribbonDropDownItemImpl1.Label = "CP932";
      ribbonDropDownItemImpl2.Label = "UTF-8";
      this.encodingComboBox.Items.Add(ribbonDropDownItemImpl1);
      this.encodingComboBox.Items.Add(ribbonDropDownItemImpl2);
      this.encodingComboBox.Label = "エンコード";
      this.encodingComboBox.Name = "encodingComboBox";
      this.encodingComboBox.Text = "UTF-8";
      // 
      // CSVEdirRibbon
      // 
      this.Name = "CSVEdirRibbon";
      this.RibbonType = "Microsoft.Excel.Workbook";
      this.Tabs.Add(this.csvEditRab);
      this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.CSVEdirRibbon_Load);
      this.csvEditRab.ResumeLayout(false);
      this.csvEditRab.PerformLayout();
      this.CSVFileGroup.ResumeLayout(false);
      this.CSVFileGroup.PerformLayout();
      this.group1.ResumeLayout(false);
      this.group1.PerformLayout();
      this.ResumeLayout(false);

    }

    #endregion

    internal Microsoft.Office.Tools.Ribbon.RibbonTab csvEditRab;
    internal Microsoft.Office.Tools.Ribbon.RibbonGroup CSVFileGroup;
    internal Microsoft.Office.Tools.Ribbon.RibbonButton openFolderButton;
    internal Microsoft.Office.Tools.Ribbon.RibbonButton saveFolderButton;
    internal Microsoft.Office.Tools.Ribbon.RibbonButton saveFolderAsButton;
    internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
    internal Microsoft.Office.Tools.Ribbon.RibbonEditBox spearatorEditBox;
    internal Microsoft.Office.Tools.Ribbon.RibbonEditBox extEditBox;
    internal Microsoft.Office.Tools.Ribbon.RibbonComboBox encodingComboBox;
  }

  partial class ThisRibbonCollection
  {
    internal CSVEdirRibbon CSVEdirRibbon
    {
      get { return this.GetRibbon<CSVEdirRibbon>(); }
    }
  }
}
