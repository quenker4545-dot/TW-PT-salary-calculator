# 台灣 PT／短期勞工實拿薪資計算機

這是一個可部署到 Streamlit Community Cloud 的 Streamlit App。用途是快速試算 PT、活動工讀、展場工讀、短期臨時工的：

- 平日薪資
- 平日加班費
- 休息日出勤薪資
- 國定假日／雙薪日薪資
- 勞保 + 就保自付額估算
- 健保自付額估算
- 所得稅預扣估算
- 預估實拿薪資
- Excel 匯出

> 注意：此工具是估算器，不是法律意見。實際薪資與保費仍以公司申報投保薪資、薪資單、勞保局、健保署與主管機關公告為準。

---

## 1. 本機執行

```bash
python -m venv .venv

# Windows PowerShell
.venv\Scripts\Activate.ps1

# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

打開瀏覽器後進入：

```text
http://localhost:8501
```

---

## 2. 部署到 Streamlit Community Cloud

### Step 1：建立 GitHub Repository

把這個資料夾的檔案上傳到 GitHub：

```text
app.py
requirements.txt
.streamlit/config.toml
README.md
.gitignore
Dockerfile
```

### Step 2：進入 Streamlit Community Cloud

到 Streamlit Community Cloud，選擇 Deploy app，連接你的 GitHub repo。

### Step 3：設定入口檔

Main file path 輸入：

```text
app.py
```

部署完成後，你會得到一個公開網址，例如：

```text
https://your-app-name.streamlit.app
```

---

## 3. iPhone 使用方式

最穩定的方法不是做原生 iOS App，而是：

1. 用 Safari 打開你的 Streamlit App 網址。
2. 點分享按鈕。
3. 選「加入主畫面」。
4. 之後就像 App 一樣從桌面打開。

---

## 4. 原生 iOS App 外殼

如果你真的要做成 iOS App，可以用 `ios_webview_shell/SalaryWebViewApp.swift`。

它只是 WKWebView 外殼，會把你的 Streamlit 網址包進 iOS App。你仍然需要：

- Mac
- Xcode
- Apple Developer 帳號或本機簽名
- App Store Connect（若要上架）

`.ipa` 不能只靠 Python 直接產生，必須經過 Apple 簽名流程。

---

## 5. 可調參數

側欄可以調整：

- 勞保 + 就保合計費率
- 勞保勞工負擔比例
- 健保費率
- 健保本人負擔比例
- 是否手動指定投保薪資
- 勞保是否按天數比例估算
- 健保是否按天數比例估算

若官方級距改變，可以直接修改 `app.py` 內的：

```python
LABOR_INSURED_GRADES = [...]
NHI_INSURED_GRADES = [...]
```

---

## 6. Docker 執行

```bash
docker build -t pt-salary-app .
docker run -p 8501:8501 pt-salary-app
```
