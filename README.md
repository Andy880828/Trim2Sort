20230529
將中文資料庫複製到這個資料夾了。
寫了將中文資料合併上去的小script。
用pip安裝了openpyxl跟jinja2。
將資料庫存成xlsx格式並把裡面的colname的第一個字母都改成大寫。
臨時將中文小script合併進了ngs.py(為了讓顏色維持)，
加入Line23的ref = "D:\\Trim2Sort_ver3.0.0\\ref_ver3_0408-3.xlsx"，
還加了Line501、524、526、528
但之後可能要改，需要再寫GUI部分並把self多擴充一個self.ref.entry部分，再把以上內容還原。

20230607
合併建好的資料庫
dblist"要合併的檔案名稱"，out 輸出庫檔名。
blastdb_aliastool -dblist "contamination_12S 12S_Osteichthyes_db" -dbtype nucl -out 12S_Osteichthyes_db -title "12S_Osteichthyes_db"