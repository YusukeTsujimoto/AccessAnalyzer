using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessAnalyzer
{
    class Program
    {
        static void Main(string[] args)
        {
            // ディレクトリのリスト
            var orgDirectories = new List<DirectoryInfo>
            {
                new DirectoryInfo(@"C:\minakuchi\AccOrg"),
                new DirectoryInfo(@"C:\minakuchi_Eco\AccOrg"),
                new DirectoryInfo(@"C:\minakuchi_seikyu\AccOrg"),
                new DirectoryInfo(@"C:\minakuchi_tablet\AccOrg"),
            };
            var tableDictionary = new Dictionary<FileInfo, IEnumerable<string>>();
            foreach (var dir in orgDirectories)
            {
                // accdbファイルの一覧を取得する
                var accessFiles = GetAccessFiles(dir);
                foreach (var file in accessFiles)
                {
                    // テーブルの名前リストを取得する
                    var tableNames = GetTableNames(file);
                    Console.WriteLine($"{file.FullName}のテーブル名を{tableNames.Count()}取得しました");
                    if (tableNames.Any())
                    {
                        // ディクショナリに追加
                        tableDictionary.Add(file, tableNames);
                    }
                }
                Console.WriteLine($"{tableDictionary.Count}ファイル取得完了");
            }
            // Excelに出力
            WriteToExcel(new FileInfo(@"C:\Users\TDICK003\Desktop\Accessテーブル名一覧.xlsx"), tableDictionary);
        }

        public static IEnumerable<FileInfo> GetAccessFiles(DirectoryInfo directory)
        {
            return directory.GetFiles("*.accdb", SearchOption.AllDirectories);
        }

        public static IEnumerable<string> GetTableNames(FileInfo accdbFile)
        {
            string connectionString =
                "Provider=Microsoft.ACE.OLEDB.12.0;" +
                "Data Source='" + accdbFile.FullName + "';";

            try
            {
                using (var connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    System.Data.DataTable table = connection.GetSchema("Tables", new string[4] { null, null, null, "TABLE" });

                    var tableNames = new List<string>();
                    foreach (DataRow row in table.Rows)
                    {
                        tableNames.Add(row["TABLE_NAME"].ToString());
                    }
                    return tableNames;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        private static void WriteToExcel(FileInfo fileInfo, IDictionary<FileInfo, IEnumerable<string>> tableDictionary)
        {
            Application xlApp = new Application();

            //Excelが開かないようにする
            xlApp.Visible = false;

            //指定したパスのExcelを起動
            var wb = xlApp.Workbooks.Open(Filename: fileInfo.FullName);

            try
            {
                //Sheetを指定
                ((Worksheet)wb.Sheets[1]).Select();
            }
            catch (Exception ex)
            {
                //Sheetがなかった場合のエラー処理

                //Appを閉じる
                wb.Close(false);
                xlApp.Quit();

                //Errorメッセージ
                Console.WriteLine(ex.Message);
                Console.ReadLine();

                //実行を終了
                Environment.Exit(0);
            }

            try
            {
                //変数宣言
                var rowCount = 1;
                //書き込む場所を指定
                var dirHeaderCell = xlApp.Cells[rowCount, 1] as Range;
                //書き込む内容
                dirHeaderCell.Value2 = "ディレクトリ名";

                var fileHeaderCell = xlApp.Cells[rowCount, 2] as Range;
                fileHeaderCell.Value2 = "ファイル名";

                var tableHeaderCell = xlApp.Cells[rowCount, 3] as Range;
                tableHeaderCell.Value2 = "テーブル名";

                rowCount++;

                foreach (var file in tableDictionary)
                {
                    foreach (var name in file.Value)
                    {
                        var dirNameCell = xlApp.Cells[rowCount, 1] as Range;
                        dirNameCell.Value2 = file.Key.DirectoryName;
                        var fileNameCell = xlApp.Cells[rowCount, 2] as Range;
                        fileNameCell.Value2 = file.Key.Name;
                        var tableNameCell = xlApp.Cells[rowCount, 3] as Range;
                        tableNameCell.Value2 = name;
                        rowCount++;
                    }
                    Console.WriteLine($"{file.Key.FullName}の書き込みが完了しました");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
            finally
            {
                //Appを閉じる
                wb.Close(true);
                xlApp.Quit();
            }
        }
    }
}
