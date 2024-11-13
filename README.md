![image](https://github.com/Andy880828/Trim2Sort/blob/9ba7b1d031eb858739b8b20be1cf14268afed621/Trim2Sort.png)
Trim2Sort設計來大量處理fastq資料(與單筆的Sanger定序資料)，需先備有:
1. Usearch。
2. Cutadapt。
3. Blastn。
4. 建置且映射好的sequence database。
5. Excel程式(用來開啟輸出檔案)，也可以用開源的其他免費試算表軟體開啟。



-----Update-----
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
