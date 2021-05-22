using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Jp.Gr.SabotageCLA6.Utils;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace CSVDEditAddin
{
  /// <summary>CSVを編集するためのEXCELAddinのリボン</summary>
  public partial class CSVEdirRibbon
  {
    private DirectoryInfo opendDirectory = null;

    #region instance variables
    CommonOpenFileDialog folderDialog = new CommonOpenFileDialog()
    {
      Title = "CSVを格納したフォルダを選択する",
      IsFolderPicker = true,
    };
    #endregion instance variables


    #region リボンロードイベント
    /// <summary>リボンロードイベント</summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void CSVEdirRibbon_Load(object sender, RibbonUIEventArgs e)
    {

    }
    #endregion リボンロードイベント

    #region フォルダを開くボタン押下イベント
    /// <summary>フォルダを開くボタン押下イベント</summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void openFolderButton_Click(object sender, RibbonControlEventArgs e)
    {
      var dialogResult = this.folderDialog.ShowDialog();
      if (dialogResult != CommonFileDialogResult.Ok)
      {
        return;
      }

      this.opendDirectory = new DirectoryInfo(this.folderDialog.FileName);

      string direName = this.opendDirectory.Name;

      var tempFile = Path.Combine(Path.GetTempPath(), direName + ".xlsx");
      var csvbook = Globals.ThisAddIn.Application.Workbooks.Add();
      csvbook.SaveAs(tempFile);
      var defaultSheet = csvbook.Worksheets[1];
      DataSet ds = new DataSet(direName);
      foreach (var csvFile in this.opendDirectory.GetFiles(this.extEditBox.Text, SearchOption.TopDirectoryOnly))
      {
        var sheet = csvbook.Worksheets.Add(Before: defaultSheet) as Worksheet;
        sheet.Name = csvFile.Name;
        sheet.Cells.NumberFormat = "@";
        sheet.Cells.Font.Name = "Mgen+ 2m regular";

        System.Data.DataTable csvata =
          CSVUtil.ReadCSVFile(csvFile.FullName, Encoding.GetEncoding(this.encodingComboBox.Text));

        for (int columnsIndex = 0; columnsIndex < csvata.Columns.Count; columnsIndex++)
        {
          sheet.Cells[1, columnsIndex + 1].Value = csvata.Columns[columnsIndex].ColumnName;
        }

        for (int rowsIndex = 0; rowsIndex < csvata.Rows.Count; rowsIndex++)
        {
          for (int columnsIndex = 0; columnsIndex < csvata.Columns.Count; columnsIndex++)
          {
            sheet.Cells[rowsIndex + 2, columnsIndex + 1].Value = csvata.Rows[rowsIndex][columnsIndex];
          }
        }

      }
    }
    #endregion フォルダを開くボタン押下イベント


    #region フォルダに上書き保存ボタン押下イベント
    /// <summary>上書き保存ボタン押下時のイベント</summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void saveFolderButton_Click(object sender, RibbonControlEventArgs e)
    {
      var csvbook = Globals.ThisAddIn.Application.ActiveWorkbook;
      for (int sheetIndex = 1; sheetIndex < csvbook.Sheets.Count; sheetIndex++)
      {
        Worksheet sheet = csvbook.Sheets[sheetIndex];

        if(this.opendDirectory == null)
        {
          var dialogResult = this.folderDialog.ShowDialog();
          if (dialogResult != CommonFileDialogResult.Ok)
          {
            return;
          }
          this.opendDirectory = new DirectoryInfo(this.folderDialog.FileName);
        }


        using (StreamWriter writer = new StreamWriter(Path.Combine(this.opendDirectory.FullName, sheet.Name)))
        {
          for (int rowIndex = 1; !String.IsNullOrEmpty(sheet.Cells[rowIndex, 1].Text); rowIndex++)
          {
            StringBuilder line = new StringBuilder();
            for (int columnsIndex = 1; !String.IsNullOrEmpty(sheet.Cells[1, columnsIndex].Text); columnsIndex++)
            {
              string value = sheet.Cells[rowIndex, columnsIndex].Text;
              value = value.Replace("\"", "\"\"");
              value = "\"" + value + "\"";
              line.Append(value);
              line.Append(",");
            }
            if(line.Length > 0)
            {
              line.Remove(line.Length - 1, 1);
            }
            writer.WriteLine(line.ToString());
          }
        }
      }
    }
    #endregion フォルダに上書き保存ボタン押下イベント


    #region フォルダ名を指定して保存ボタン押下イベント
    /// <summary>フォルダ名を指定して保存ボタン押下イベント</summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void saveFolderAsButton_Click(object sender, RibbonControlEventArgs e)
    {
      var dialogResult = this.folderDialog.ShowDialog();
      if (dialogResult != CommonFileDialogResult.Ok)
      {
        return;
      }
      this.opendDirectory = new DirectoryInfo(this.folderDialog.FileName);

      this.saveFolderButton_Click(sender, e);
    }
    #endregion フォルダ名を指定して保存ボタン押下イベント
  }
}
