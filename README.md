# BackToTheA1
ExcelファイルのA1セルを選択状態にします。  
一番左側のシートを選択状態にします。  
各シートのA1セルを選択状態にし、倍率を100%にします。  
  
※ExcelがインストールされたPC上でのみ動作します。拡張子はxlsxのみ対応しています。

## 使用例
BackToTheA1.exe "c:\temp\hoge.xlsx"(フルパス指定のこと)  
もしくは、BackToTheA1.exeにファイルをドラッグ＆ドロップすれば変換されます。  
  
※変換後、元には戻せませんので、予めバックアップを取得しておいてから実行してください！

## 戻り値
0:正常終了（変換成功）  
1:異常終了（変換失敗。エラー詳細はコンソール出力を参照）

## 使用ライブラリ
NetOffice.Excel.Net45.1.7.4.4  
NetOffice.Core.Net45.1.7.4.4
（何れもNuGetから取得しています）

## ライセンス
MITライセンス

## 作成
https://juraku-software.net/
  
コンパイルしてすぐに使える状態のEXEは以下のページで公開しています。  
https://juraku-software.net/windows-app-backtothea1/
