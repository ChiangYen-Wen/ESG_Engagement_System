# ESG 議合報告上傳系統_投信專案
## ESG Engagement Report Management System

> 可執行於 **Flask + Access** 資料庫上的簡易網站應用，用來管理 ESG 報告，包括「查詢」、「新增」、「修改」和「刪除」功能。

> 本系統定位為簡易快速使用，符合主管機關要求內部管理需求。

✅ **特別強調：目前已經打包成獨立的 .exe 檔案，可直接點選連結下載範例**

✅ **不需要安裝 Python 環境：可直接在企業內部 Windows 電腦上運行**

✅ **範例連結：https://drive.google.com/drive/folders/1dNrKkeXgRzzUz3KOg2xZah6Sdtdcdc6q?usp=drive_link**

---

## 👨‍💼 簡易使用說明：

> 這個系統讓內部同仁不需要懂技術，只要打開應用程式 (.exe)，在網頁介面上就能輕鬆完成 ESG 報告的新增、查詢與管理，快速提升內部資料整理與分析效率！
>
> ✅ **免安裝**、✅ **快速啟動**、✅ **降低訓練成本**

## 📥 詳細部署安裝使用說明

### 資料夾結構與存取

- `app.py` — 後端程式，位於專案根目錄
- `templates/index.html` — 主頁（查詢＋新增），放在 `templates` 子資料夾
- `templates/detail.html` — 報告詳細內容頁面
- `templates/edit.html` — 報告編輯頁面
- `ESG_DB.accdb` — Access 資料庫檔案，放在專案根目錄

  
**注意**：
- HTML 檔案必須放在 `templates/` 目錄內，Flask 才能正確載入。
- **`ESG_DB.accdb` 必須自備，並且和 `app.py` 放在同一層資料夾**，否則系統無法連線資料庫。
- `外層資料夾截圖畫面如下:
- ![螢幕擷取畫面 2025-04-26 190502](https://github.com/user-attachments/assets/9fb9e076-da1e-4817-bc60-ae371dc3bde0)
- `內層資料夾截圖畫面如下:
- ![螢幕擷取畫面 2025-04-26 190247](https://github.com/user-attachments/assets/f718cf43-9094-4b74-ab5e-f5db04af09ee)


---

## 🚿 安裝與啟動

### 1. 環境要求（原始碼版）

- Python 3.8+
- Flask
- pyodbc 套件
- Microsoft Access ODBC Driver

安裝套件：

```bash
pip install flask pyodbc
```

### 2. 啟動步驟（使用原始碼）

確保 `app.py` 與 `ESG_DB.accdb` 同資料夾，依序執行：

```bash
python app.py
```

會自動打開瀏覽器，可以有操作介面直接使用(**內容僅限本地端可存取**)

### 3. 啟動步驟（使用打包成 .exe 版本）

直接點擊 `.exe` 檔案即可啟動，不需要安裝任何環境！

### 4. 如何自行打包成 .exe？

若需要自行打包，請按照以下步驟：

1. 安裝 pyinstaller：

```bash
pip install pyinstaller
```

2. 在專案根目錄執行打包指令：

```bash
pyinstaller --add-data "templates;templates" --add-data "ESG_DB.accdb;." --onefile app.py
```

- `--add-data` 是指定要包含的資料夾與檔案
- `--onefile` 表示打包成單一可執行檔
- 打包後的 `.exe` 檔案會出現在 `dist/` 資料夾內

**注意：** 因為打包後的 `.exe` 檔案體積通常會超過 25MB，無法直接上傳至 GitHub。

✅ 解決方法：
- 將 `.exe` 檔案上傳到 Google Drive、Dropbox、OneDrive 等雲端硬碟，並在 GitHub ReadMe 中提供下載連結。
- 或者，壓縮成 .zip 檔案後，使用 GitHub Release 功能發佈大檔案。

✅ 具體流程圖如下：
![ChatGPT Image 2025年4月26日 下午07_21_27](https://github.com/user-attachments/assets/91aa5991-ae76-4657-917b-df6dedd02759)


---

## 🔍 功能簡介

### 📅 ESG 報告查詢

- 根據股票代號、分析師、目的、日期篩選
- 結果列表，點擊可查看詳細內容
「查詢」、「新增」、「修改」和「刪除」
### 📄 新增 ESG 報告

- 填寫表單，直接送出並存入 Access 資料庫

### ✏️ 修改 ESG 報告

- 顯示完整的報告資訊（日期、股票、問答內容等）

### 🗑️ 刪除資訊

- 可選定特定的ESG報告進行刪除

---

## ⚠️ 注意事項

- `ESG_DB.accdb` 資料庫必須包含 **`ESG_report`** 這個表格，欄位名稱需與程式對應
- 系統尚未進行安全性強化（例如 CSRF 防禦、輸入驗證），僅建議內部使用

---



歡迎 Fork 與 PR，一起使系統更強大！🚀


