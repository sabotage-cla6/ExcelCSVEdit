using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jp.Gr.SabotageCLA6.Utils
{
  /// <summary>CSV関連のUtilクラス</summary>
  public static class CSVUtil
  {
    /// <summary>CSVファイルを読み込んで、DataTableに展開します。</summary>
    /// <param name="filePath">読み込みCSVのファイル</param>
    /// <param name="encoding">エンコード</param>
    /// <returns>展開結果のDataTable</returns>
    public static DataTable ReadCSVFile(string filePath, Encoding encoding)
    {
      DataTable reslut = new DataTable(Path.GetFileName(filePath));

      int columnsIndex = 0;
      using (StreamReader reader = new StreamReader(filePath, encoding))
      {
        // CSVヘッダの読み込み
        ReadCSVRecord(reader, val =>
        {
          reslut.Columns.Add(val);
        });

        // CSVデータ行の読み込み
        while (!reader.EndOfStream)
        {
          DataRow addRow = reslut.Rows.Add();
          columnsIndex = 0;
          ReadCSVRecord(reader, val =>
          {
            addRow[columnsIndex++] = val;
          });
        }
      }

      return reslut;
    }

    /// <summary>CSVを１レコード読み込みます</summary>
    /// <param name="reader">CSVのStreamReader</param>
    /// <param name="setValueAction">各要素の読み取り結果を記録するためのアクション</param>
    private static void ReadCSVRecord(StreamReader reader, Action<string> setValueAction)
    {
      bool isWaitEndQuote = false;
      StringBuilder element = new StringBuilder();
      while (!reader.EndOfStream)
      {
        string text = reader.ReadLine() + '\n';
        char[] cArray = text.ToCharArray();
        for (int i = 0; i < cArray.Length; i++)
        {
          char c = cArray[i];

          // "..[c] の状態で読み込み中
          if (isWaitEndQuote)
          {
            switch (c)
            {
              case '"':
                if (i + 1 < cArray.Length && cArray[i + 1] == '"')
                {
                  i++;
                  element.Append(c);
                }
                else
                {
                  isWaitEndQuote = false;
                }
                break;
              default:
                element.Append(c);
                break;
            }
          }
          // ..[c] の状態で読み込み中
          else
          {
            switch (c)
            {
              case '"':
                if (element.Length == 0)
                  isWaitEndQuote = true;
                break;
              case ',':
                setValueAction.Invoke(element.ToString());
                element.Clear();
                break;
              case '\r':
              case '\n':
                setValueAction.Invoke(element.ToString());
                element.Clear();
                return;
              default:
                element.Append(c);
                break;
            }
          }
        }
      }
    }
  }
}
