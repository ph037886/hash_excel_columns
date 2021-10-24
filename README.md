# hash_excel_columns
hash one excel file with sheets and columns

這個code是在要把資料給單位實習生的時候寫的小專案
主要是要將單位內的Excel部分欄位做hash之後再轉出
一開始想得很簡單，覺得靠迴圈就能搞定
不過開始做第一個檔案就發現許多問題，主要問題是一個excel裡面有不同sheet，然後一個sheet裡面也不只一個欄位需要hash
所以就總結出這個小專案

1. 包含讀入與存的視窗
2. 讀取讀入檔名，並作為預設存檔檔名
3. 每一個sheet分別執行
4. 輸入欄位號碼
5. 用亂數產生salt，同一個Excel裡面用同一個salt，這樣同時有去辨識也能在同一個excel中做分組運算
