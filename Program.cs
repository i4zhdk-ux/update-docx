using System;
using System.Diagnostics;
using System.IO;
using NPOI.XWPF.UserModel;
using System.Globalization;

namespace WordTableReader
{
    class Program
    {
        static void Main(string[] args)
        {
            // 1. 引数のチェック
            if (args.Length == 0)
            {
                Console.WriteLine("使用法: wordApp.exe <ファイルパス>");
                return;
            }

            string filePath = args[0];

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"エラー: ファイルが見つかりません: {filePath}");
                return;
            }

            // 2. ストップウォッチの開始
            Stopwatch sw = new Stopwatch();
            sw.Start();

            try
            {
                Console.WriteLine($"{Path.GetFileName(filePath)} を読み取り中...");

                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    // Wordドキュメントの読み込み
                    XWPFDocument doc = new XWPFDocument(fs);

                    // 3. 表（テーブル）の読み取り
                    int tableCount = 1;
                    foreach (var table in doc.Tables)
                    {
                        if(table.Rows.Count < 2)
                        {
                            continue;
                        }
                        Console.WriteLine($"{tableCount++}ページ目を読み取り中");
                        for (int i = 2; i < table.Rows.Count; i++)
                        {
                            var row = table.Rows[i];
                            if (i % 2 == 0) 
                            {
                                continue;
                            }
                            
                            var cells = row.GetTableCells();
                            for (int j = 2; j < cells.Count; j++)
                            {
                                string cellValue = cells[j].GetText();
                                var preRow = table.Rows[i - 1];
                                var preCells = preRow.GetTableCells();
                                if (preCells[0].GetText() == "")
                                {
                                    break;
                                }
                                if (j == 2 && cellValue == "") {
                                    break; 
                                } else if (j % 2 == 1 && cellValue == "") {
                                    var name = preCells[j - 1].GetText();
                                    preCells[2].Paragraphs[0].Runs[0].SetText(name, 0);
                                    var date = cells[j - 1].GetText();
                                    cells[2].Paragraphs[0].Runs[0].SetText(date, 0);
                                    for (int k = 3; k < j; k++)
                                    {
                                        cells[k].Paragraphs[0].Runs[0].SetText("", 0);
                                        if (k % 2 == 0)
                                        {
                                            table.Rows[i - 1].GetTableCells()[k].Paragraphs[0].Runs[0].SetText("", 0);  
                                        }                                        
                                    }
                                    break;
                                } else if (cellValue == "") {
                                    var date = cells[j - 1].GetText();
                                    preCells[1].Paragraphs[0].Runs[0].SetText(date, 0);
                                    for (int k = 2; k < j; k++)
                                    {
                                        cells[k].Paragraphs[0].Runs[0].SetText("", 0);
                                        if (k % 2 == 0)
                                        {
                                            table.Rows[i - 1].GetTableCells()[k].Paragraphs[0].Runs[0].SetText("", 0);  
                                        }                                        
                                    }
                                    break;
                                }
                            }
                        }
                    }

                    using (FileStream fsw = new FileStream("更新結果.docx", FileMode.Create, FileAccess.Write))
                    {
                        doc.Write(fsw);
                    }
                }

                // 4. 計測終了
                sw.Stop();
                Console.WriteLine("----------------------------------");
                Console.WriteLine($"処理完了。");
                Console.WriteLine($"実行時間: {sw.ElapsedMilliseconds} ms ({sw.Elapsed.TotalSeconds:F2} 秒)");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"エラーが発生しました: {ex.Message}");
            }
        }
    }
}