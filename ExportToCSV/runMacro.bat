@echo off

echo 処理開始

rem VBScript呼び出し 第一引数:マクロの入ったExcelファイルパス 第二引数:モジュール名.マクロ名
cscript action.vbs "C:\Users\kobat88\Desktop\VBA\ExportToCSV\TT69K-E3.xlsm" "ExportToCSV.ExportToCSV"

rem 以下でエラー処理するには､vbsからWScript.Quit(x)でerrorLevelを返す必要あり。
rem if %errorLevel% neq 0 (
rem	echo;
rem	echo エラー発生
rem	echo errorLevelは%errorLevel%です。
rem	pause
rem ) else (
rem	echo;
rem	echo 正常終了
rem	pause
rem )

echo 処理終了
pause
