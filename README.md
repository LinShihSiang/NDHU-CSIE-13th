# NDHU-CSIE-13th
計算畢旅各人收支程式

## 使用方式
1. 請將輸入 xlsx 檔放入 Input 資料夾
2. 修改 appsettings.json FileName 與放入的檔名一致
3. 重建後執行 NDHU-CSIE-13th.exe
4. 結算結果將產於 Output 資料夾

## Input
請參考專案底下 Input/NDHU-CSIE#13 - 10th.xlsx 範本進行設定
總共分為 3 個 Sheet
1. 個人代墊款項
2. 畢旅開銷明細
3. 參與人員及所參與之活動明細

## Output
檔名 {InputFileName}_結算.xlsx，並包含以下 Sheet
1. 個人代墊款項
2. 畢旅開銷明細 (包含參與人員總數、每人平均金額)
3. 參與人員各別 Sheet (包含人員參與項目、代墊項目、應給付、應付款等資訊)
