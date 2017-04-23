下載股市資料:
參數1:Stock.exe檔案位置 
參數2:Download
參數3:查詢年(最早94/09/01開始)
參數4:查詢月
參數5:查詢日
參數6:分類項目代號
參數7:輸出檔案位置(最後要有\)
範例:Stock.exe Download 106 04 07 ALL D:\StockOutputData\Source\

分析股市資料:
參數1:Stock.exe檔案位置 
參數2:Analyze
參數3:來源資料夾
參數4:目標資料夾(最後要有\)
範例:Stock.exe Analyze D:\StockOutputData\Source D:\StockOutputData\

注意事項:
1.只處理來源資料夾內的檔案(不包含子層檔案)
2.當月份寫入到目標資料夾的資料都是重新計算覆寫的，不會累計
3.Backup資料夾內的AnalysisData(初始資料).xls為最初的空資料，備份用
4.每次執行完股市分析後，記得備份到Backup資料夾內，若之後有問題才不用再全部重做
5.資料遇到"-"或"0.00"皆不會列入計算，#NUM!表示除以0的結果
6.Exe資料夾放的是程式的執行檔