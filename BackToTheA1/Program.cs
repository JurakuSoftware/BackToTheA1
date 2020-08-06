using System;
using System.IO;
using Excel = NetOffice.ExcelApi;

namespace BackToTheA1
{
    /// <summary>
    /// 指定されたExcelファイルの選択セルをA1セルに変更する
    /// ・一番左側のシートを選択状態にする
    /// ・各シートの選択セルをA1にする
    /// ・各シートの倍率を100%にする
    /// 
    /// プログラム戻り値
    /// 0 = 正常終了（変換完了）
    /// 1 = 異常終了（何らかのエラー発生。エラー内容はコンソール出力を参照）
    /// </summary>
    public class Program
    {
        /// <summary>
        /// プログラム戻り値（正常終了）
        /// </summary>
        private const int OK = 0;
        /// <summary>
        /// プログラム戻り値（異常終了）
        /// </summary>
        private const int NG = 1;


        /// <summary>
        /// 引数１で指定されたファイルパス
        /// </summary>
        private static string filePath = "";

        /// <summary>
        /// ExcelファイルのA1セル選択処理
        /// 引数１：Excelファイルパス
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        public static int Main(string[] args)
        {
            //引数チェック
            if (!CheckArgs(args)) return NG;

            //Excel起動
            WriteLog("Excel起動待ち");
            using (Excel.Application excelApplication = new Excel.Application())
            {
                WriteLog("Excel起動");

                try
                {
                    //Excel描画停止（速度改善のため）
                    ExcelBeginUpdate(excelApplication);

                    //対象ファイルを開く
                    WriteLog("ファイル開く");
                    Excel.Workbook workBook = excelApplication.Workbooks.Open(filePath);

                    //A1セルを選択状態にする
                    WriteLog($"全シート数：{workBook.Sheets.Count}");
                    for (int i = 1; i <= workBook.Sheets.Count; i++)
                    {
                        WriteLog($"{i}シート目処理中");

                        //シート取得
                        Excel.Worksheet sheet = (Excel.Worksheet)workBook.Sheets[i];
                        WriteLog($"シート名：{sheet.Name}");

                        //非表示のシートは操作に失敗するので無視する
                        if (sheet.Visible != Excel.Enums.XlSheetVisibility.xlSheetVisible) continue;

                        //シートを選択状態にする（こうしないとセル選択に失敗する）
                        sheet.Select();

                        //倍率を100%に変更
                        excelApplication.ActiveWindow.Zoom = 100;

                        //一番左上のA1セルを選択状態に変更
                        sheet.Range("A1").Select();
                    }

                    //一番左に存在するシートを選択状態にする
                    WriteLog("一番左のシートを選択");
                    Excel.Worksheet firstSheet = (Excel.Worksheet)workBook.Sheets[1];
                    firstSheet.Select();

                    //保存
                    WriteLog("保存");
                    workBook.Save();
                    workBook.Close();

                    //Excel終了
                    excelApplication.Quit();
                }
                catch (Exception ex)
                {
                    //何らかのエラー発生
                    WriteLog("例外エラー発生", ex);
                    return NG;
                }
                finally
                {
                    //Excelを閉じる
                    try
                    {
                        excelApplication.Quit();
                        excelApplication.Dispose();
                    }
                    catch { } //例外発生時は無視
                }
            }

            return OK;
        }

        /// <summary>
        /// 起動引数チェック
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        private static bool CheckArgs(string[] args)
        {
            //引数未指定チェック
            if (args == null || args.Length == 0) { WriteLog("ファイル名を指定してください"); return false; }

            //ファイル存在チェック
            filePath = args[0];
            if (!File.Exists(filePath)) { WriteLog("指定されたファイルは存在しません。パスに誤りが無いかをご確認ください"); return false; }

            //拡張子チェック
            if (Path.GetExtension(filePath).ToLower() != ".xlsx") { WriteLog("拡張子はxlsxのみ対応しています"); return false; }

            //ファイル使用中チェック
            if (CheckFileInUse()) { WriteLog("指定されたファイルは使用中のため、変換出来ません"); return false; }

            return true;
        }

        /// <summary>
        /// ファイル使用中チェック
        /// </summary>
        /// <returns></returns>
        public static bool CheckFileInUse()
        {
            try
            {
                //移動元と移動先を同じにすることで、何も変化させない
                //ただし、Excelやサクラエディタなどで開いている場合は、例外エラーが発生するためこれを利用してファイル使用中と判断する
                File.Move(filePath, filePath);
            }
            catch (IOException)
            {
                //ファイル使用中
                return true;
            }

            //ファイル使用中でない
            return false;
        }

        /// <summary>
        /// Excel描画停止
        /// </summary>
        private static void ExcelBeginUpdate(Excel.Application excelApplication)
        {
            //処理速度を上げるために、描画停止
            excelApplication.ScreenUpdating = false;
            excelApplication.Visible = false;

            //ダイアログを表示しない
            excelApplication.DisplayAlerts = false;
        }

        /// <summary>
        /// ログ出力（コンソールへ）
        /// </summary>
        /// <param name="s"></param>
        /// <param name="ex"></param>
        private static void WriteLog(string s, Exception ex = null)
        {
            //例外エラー発生時は、その内容も記載
            if (ex != null) s = $"[ERROR]{s}\r\n{ex.ToString()}";

            //コンソール出力
            Console.WriteLine(s);
        }
    }
}
