【ezPAY 自動化生成批次開立發票 csv 檔案小工具】
​
我們公司使用的電子發票平台是 ezPAY，如果有使用過 ezPAY 批次開立的朋友，應該會知道需要上傳一個超級無比複雜的 csv 檔。我真的不知道其他人是怎麼批次開立的，我真的是會被搞死。
​
好的，所以再次使用 google spreadsheet 加上 AppScript 腳本，讓我可以增加這次活動的資料之後，自動產生對應格式的上傳檔案，已經實驗成功了，分享給大家。
​
▌使用方法：
1. 複製一份試算表到你的雲端 https://docs.google.com/spreadsheets/d/1OJ81X5Bp3A7Rjg37wMCCP4M92kHeXdGwSxMc2GU_AI8/edit?usp=sharing
2. 把工作表改成要上傳檔案那天的日期，例如你預計 20240525 上傳，那你就把工作表改成 20240525
3. 開始修改腳本
- 打開你複製的 Google 試算表。
- 點擊上方選項中的擴充功能，選擇 Apps Script。
- 把腳本裡面要改成你的公司相關資料的地方替換掉，或是請 gpt 幫你改，以及增加一些我現在沒有，但你需要的欄位，例如買受人地址（prompt 可參考圖片）
 
// 表頭記錄 這裡要改 ezPay電子發票會員編號、商店代號 這兩個
 ​ let csvContent = 'H,INVO,C12317769563,326347694,' + sheetName + ',,,,,,,,,,,\n';
​
// 同一區的最下面我有設定生成的檔案檔名，這裡也要改成你的商店代號
 ​ const fileName = '326347694_' + sheetName + '.csv';
 ​ const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
​
 // 指定文件夾 ID 後面改成你雲端後面的網址 https://drive.google.com/drive/u/1/folders/ 這串後面的網址
 ​ const folderId = '1hwnEhYitCMIdaWUkzr6uFrj89eDCm2';
​
- 點上面的儲存專案。
​
4. 重新載入試算表，或者在 Apps Script 編輯器中運行 onOpen 函數。（或者後面那句是 gpt 加的我看不懂）
5. 在上方選項中會多一個 Custom Menu 可以選擇 "Generate Invoice CSV" 項目來運行腳本。
6. 他會有要你授權執行，然後會有很警示的存取權，如果被擋住了可以點選進階，然後接受使用，風險就大家自己斟酌要不要用～
7. 開立成功之後，工作表最後面會多一欄執行時間，還有檔案所在位置。下載檔案就可以去上傳 ezPAY 了！
​
▌功能說明
1. 會依照你頁面所在的工作表執行，每張工作表是一個日期，建議同一個日期的發票都一起開立，才不會遇到流水號重複的問題。
2.這個檔案有些資料是空白，像是買受人地址、載具類別、載具編號、捐贈碼等欄位，如果有要帶入這些資料，就再跟 gpt 溝通你要的 S: 銷售明細 每一欄是什麼（以上就是 ghij 這四欄的內容），並增加 input 的表頭。Appscript 的程式碼跟表頭要一起增改才不會出錯。
​
▌踩到的雷
1. S: 銷售明細 多空了一欄 K 的問題，在使用 gpt 提供的腳本產生 csv 後，發現 S 這一列的空白多空了一欄，我好說歹說他都刪不掉，我只好手動改 Appscript 裡面的程式碼。

▌gpt 協作過程（可能失效打不開）
https://chatgpt.com/share/e/49cf0615-cece-4a6a-829e-80fb6785ed31
