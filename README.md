# PDF 列管表 → Google 試算表

把資料夾裡的會議列管表 PDF，自動整理成 Google 試算表：

| 列管案號 | 主席裁示 | 承辦單位 | 最新辦理情形 | 列管建議 | 來源 PDF 名稱 |
|---|---|---|---|---|---|

處理流程：

1. **解析 PDF**：以 [docling](https://github.com/docling-project/docling) 把每份 PDF 轉成 JSON（存在 `<PDF 資料夾>/_docling_json/`）。
2. **整理資料**：
   - 合併跨頁斷開的表格。
   - 依 PDF 原本的換行分段。
   - 半形標點轉全形，中文與英數字之間加空格。
3. **寫入試算表**：資料送到 Google Apps Script，由它寫入試算表，不需要 Google Cloud 帳號或金鑰。

---

## 目錄

- [檔案說明](#檔案說明)
- [一、設定 Google 試算表（Mac、Windows 相同，只需做一次）](#一設定-google-試算表macwindows-相同只需做一次)
- [快速安裝：用 pipx 安裝成 pdf2sheet 指令（推薦）](#快速安裝用-pipx-安裝成-pdf2sheet-指令推薦)
- [二、Mac 安裝與使用](#二mac-安裝與使用)
- [三、Windows 安裝與使用](#三windows-安裝與使用)
- [四、指令選項](#四指令選項)
- [五、資料整理規則](#五資料整理規則)
- [六、注意事項](#六注意事項)
- [七、常見問題](#七常見問題)

---

## 檔案說明

| 檔案 | 用途 |
|---|---|
| `pdf_to_sheet.py` | 在電腦上執行：解析 PDF、整理資料、上傳 |
| `Code.gs` | 貼進 Google 試算表的 Apps Script：接收資料並寫入 |

---

## 一、設定 Google 試算表（Mac、Windows 相同，只需做一次）

1. 開啟要寫入的 Google 試算表，點選「**擴充功能**」→「**Apps Script**」。
2. 刪除編輯器裡原有的程式碼，貼上 `Code.gs` 的全部內容。
3. 修改第一段的密碼，之後執行時要用同一組：
   ```js
   const TOKEN = 'change-me';   // 改成你自己的密碼
   ```
4. 點選「**部署**」→「**新增部署作業**」→ 齒輪圖示選「**網頁應用程式**」：
   - 執行身分：**我**
   - 誰可以存取：**任何人**
5. 依畫面完成授權。如果出現「Google 尚未驗證這個應用程式」，請點「進階」→「前往（不安全）」。這是你自己的程式，可以放心授權。
6. 複製「**網頁應用程式網址**」，格式為 `https://script.google.com/macros/s/…/exec`。
7. 用瀏覽器打開這個網址。看到 `{"ok":true,"message":"列管表接收端運作中"}`，就代表設定成功。

> **為什麼選「任何人」？** Python 程式不會用你的 Google 帳號登入，所以網址必須開放呼叫。密碼不對的請求一律會被拒絕，**網址與密碼請勿外流**。

> **修改 `Code.gs` 後要重新部署**：「部署」→「管理部署作業」→ 鉛筆圖示 → 版本選「**新版本**」→「部署」。網址不會改變。

---

## 快速安裝：用 pipx 安裝成 `pdf2sheet` 指令（推薦）

用 pipx 安裝後，在任何資料夾都能直接輸入 `pdf2sheet`，不用手動建立或啟用虛擬環境。專案資料夾要包含 `pyproject.toml`、`pdf_to_sheet.py`、`README.md` 三個檔案。

**Mac**

```bash
brew install pipx python@3.12
pipx ensurepath          # 執行後關閉再重新打開終端機
pipx install --editable --python /opt/homebrew/bin/python3.12 ~/Documents/Programing/PDFtoSheet
# 需要 OCR 時：
pipx inject pdf2sheet ocrmac
```

**Windows（PowerShell）**

```powershell
winget install Python.Python.3.12      # 已安裝可略過
py -3.12 -m pip install --user pipx
py -3.12 -m pipx ensurepath             # 執行後關閉再重新打開 PowerShell
pipx install --editable --python 3.12 "$HOME\Documents\PDFtoSheet"
```

**使用**

```bash
pdf2sheet ~/Desktop/pdfs --dry-run
pdf2sheet ~/Desktop/pdfs --webapp "https://script.google.com/macros/s/…/exec" --token "你的密碼"
```

指令選項與下方「四、指令選項」相同。網址與密碼也可以用環境變數設定（見二、三節的第 4 步）。

**更新、移除**

- 修改 `pdf_to_sheet.py`：因為用了 `--editable`，存檔後直接生效。
- 修改 `pyproject.toml`（例如新增套件）：`pipx install --force --editable --python <同上> <專案路徑>`
- 更新 docling：`pipx upgrade-all --include-injected` 或 `pipx runpip pdf2sheet install -U docling`
- 移除：`pipx uninstall pdf2sheet`

---

## 二、Mac 安裝與使用

### 1. 安裝 Python 3.12

建議使用 **Python 3.12**。docling 依賴的 PyTorch 對最新版 Python（例如 3.15）的支援通常比較晚。

使用 [Homebrew](https://brew.sh/)：

```bash
brew install python@3.12
```

### 2. 建立專案資料夾與虛擬環境

```bash
mkdir -p ~/Documents/PDFtoSheet
cd ~/Documents/PDFtoSheet
# 把 pdf_to_sheet.py 放進這個資料夾

/opt/homebrew/opt/python@3.12/bin/python3.12 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install docling
```

> Intel 晶片的 Mac 請把路徑改為 `/usr/local/opt/python@3.12/bin/python3.12`。

如果需要處理**掃描型 PDF**（OCR），請另外安裝：

```bash
python -m pip install ocrmac
```

### 3. 執行

每次打開新的終端機視窗，都要先啟用虛擬環境（提示字元前面會出現 `(.venv)`）：

```bash
cd ~/Documents/PDFtoSheet
source .venv/bin/activate
```

先試跑，不上傳，確認整理結果：

```bash
python pdf_to_sheet.py ~/Desktop/pdfs --dry-run
open ~/Desktop/pdfs/_docling_json/structured.csv
```

確認沒問題後，寫入試算表：

```bash
python pdf_to_sheet.py ~/Desktop/pdfs \
  --webapp "https://script.google.com/macros/s/…/exec" \
  --token "你的密碼"
```

### 4. 省略網址與密碼（選用）

把網址和密碼寫進 `~/.zshrc`，之後執行就不用每次輸入：

```bash
echo 'export GAS_WEBAPP_URL="https://script.google.com/macros/s/…/exec"' >> ~/.zshrc
echo 'export GAS_TOKEN="你的密碼"' >> ~/.zshrc
source ~/.zshrc
```

之後只要執行：

```bash
python pdf_to_sheet.py ~/Desktop/pdfs
```

---

## 三、Windows 安裝與使用

以下指令都在 **PowerShell** 中執行（開始選單搜尋「PowerShell」）。

### 1. 安裝 Python 3.12

方法 A：用 winget（Windows 10 / 11 內建）

```powershell
winget install Python.Python.3.12
```

方法 B：到 [python.org](https://www.python.org/downloads/windows/) 下載 3.12 安裝程式。安裝時**務必勾選「Add python.exe to PATH」**。

安裝完成後，**關閉再重新打開 PowerShell**，並確認版本：

```powershell
py -3.12 --version
```

### 2. 建立專案資料夾與虛擬環境

```powershell
mkdir $HOME\Documents\PDFtoSheet
cd $HOME\Documents\PDFtoSheet
# 把 pdf_to_sheet.py 放進這個資料夾

py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install docling
```

> 如果 `Activate.ps1` 出現「**因為這個系統上已停用指令碼執行**」，請先執行下面這行，再重新啟用虛擬環境：
> ```powershell
> Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
> ```

> Windows 沒有 Mac 內建的文字辨識。使用 `--ocr` 時會自動改用 EasyOCR，第一次使用會再下載 OCR 模型。

### 3. 執行

每次打開新的 PowerShell 視窗，都要先啟用虛擬環境：

```powershell
cd $HOME\Documents\PDFtoSheet
.\.venv\Scripts\Activate.ps1
```

先試跑，不上傳：

```powershell
python pdf_to_sheet.py "$HOME\Desktop\pdfs" --dry-run
start "$HOME\Desktop\pdfs\_docling_json\structured.csv"
```

寫入試算表。PowerShell 的換行符號是反引號 `` ` ``：

```powershell
python pdf_to_sheet.py "$HOME\Desktop\pdfs" `
  --webapp "https://script.google.com/macros/s/…/exec" `
  --token "你的密碼"
```

### 4. 省略網址與密碼（選用）

```powershell
setx GAS_WEBAPP_URL "https://script.google.com/macros/s/…/exec"
setx GAS_TOKEN "你的密碼"
```

設定後要**重新打開 PowerShell** 才會生效。之後只要執行：

```powershell
python pdf_to_sheet.py "$HOME\Desktop\pdfs"
```

---

## 四、指令選項

```
python pdf_to_sheet.py <PDF 資料夾> [選項]
```

| 選項 | 說明 |
|---|---|
| `--webapp URL` | Apps Script 網址（`…/exec`），或設環境變數 `GAS_WEBAPP_URL` |
| `--token 密碼` | 與 `Code.gs` 中 `TOKEN` 相同，或設環境變數 `GAS_TOKEN` |
| `--worksheet 名稱` | 寫入的工作表，預設「列管表」，不存在會自動建立 |
| `--mode upsert` | **預設**。同一份 PDF 的同一案號會覆寫原本那列，其餘新增 |
| `--mode append` | 全部新增在最後面，不檢查重複 |
| `--mode replace` | 清空工作表後重新寫入 |
| `--dry-run` | 不上傳，只輸出 `structured.csv`、`structured.json` 供檢查 |
| `--force` | 即使已有 JSON 也重新解析 PDF（修改 PDF 或更新程式後使用） |
| `--from-json` | 跳過 PDF 解析，直接讀 `_docling_json` 內現有的 JSON |
| `--json-dir 路徑` | 指定 JSON 與檢查檔的輸出位置 |
| `--recursive` | 一併處理子資料夾內的 PDF |
| `--ocr` | 掃描型 PDF（圖片檔）啟用文字辨識 |
| `-v` | 顯示詳細處理過程 |

輸出的檔案（都在 `<PDF 資料夾>/_docling_json/`）：

| 檔案 | 內容 |
|---|---|
| `<PDF 名稱>.json` | docling 原始解析結果；PDF 沒改動就不會重新解析 |
| `structured.csv` | 整理後的表格，可用 Excel / Numbers 開啟檢查 |
| `structured.json` | 同上，另含每筆案件來自哪幾頁（`_pages`） |

---

## 五、資料整理規則

### 跨頁表格合併

- PDF 每頁的表格會依序接起來，每頁重複出現的表頭會自動略過。
- 新頁開頭若是**序號與案號都空白**的列，就是上一筆案件的延續，會合併回同一格。
- 雙列表頭（例如「列管」／「案號」）會自動上下串接後辨識，不依賴固定欄位順序。
- 如果某筆案件跨頁，執行時會顯示「案號 xxx 跨頁合併（第 2、3、4 頁）」。

### 分行

- 回到原始 PDF，依每一行的實際位置判斷：
  - 該行**沒有寫到儲存格右緣**，代表段落結束，保留換行。
  - 該行**寫滿**，代表只是自動折行，與下一行接起來。
- 條列符號一律另起新行：`一、`、`（一）`、`1.`、`1、`、`（1）`、`①`。
- 「115 年 4 月 9 日列管事項」這類日期小標題，前後都會換行。

### 標點與空格

- 半形標點轉全形：`, ; : ? ! ( ) [ ] ~` → `， ； ： ？ ！ （ ） ［ ］ ～`，`"…"` → `「…」`。
- 中文後的 `.` 轉為 `。`，條列編號 `1.` 轉為 `1．`。
- **不轉換的情況**：千分位 `2,383`、小數 `83.5`、時間 `10:30`、網址、英文縮寫。
- 中文與英數字之間加一個空格，例如 `115年4月` → `115 年 4 月`；緊鄰標點符號時不加，例如 `第 4 條，`。

### 欄位格式

- **列管案號**：純數字（`1150409`）存成數字；含連字號（`1140905-2`）保留文字，避免被試算表誤判成日期。
- **列管建議**：`█除管 □續管` 只保留勾選的項目，即「除管」。
- **承辦單位**：合併成一行。

---

## 六、注意事項

- **首次執行較慢**：docling 會下載版面分析與表格結構模型，總計約數百 MB，存在 `~/.cache/huggingface/`（Windows 為 `C:\Users\<你>\.cache\huggingface\`）。之後會直接讀取，不會重複下載。模型都在本機運算，**PDF 內容不會上傳到 docling 或其他 AI 服務**，只有整理好的表格資料會送到你的 Google 試算表。
- **先用 `--dry-run` 檢查**：第一次處理新格式的 PDF 時，建議先確認 `structured.csv` 再上傳。
- **預設 `upsert` 可重複執行**：比對「列管案號 + 來源 PDF 名稱」，同一份 PDF 跑兩次不會產生重複列。換了 PDF 檔名就會被視為新資料。
- **`append` 可能重複**：網路逾時時程式會自動重試，若第一次其實已寫入，就會多出重複列。
- **掃描型 PDF**：加 `--ocr`。無論有沒有 OCR，換行判斷都需要 PDF 有文字層；沒有文字層時改以文字規則分段，準確度較低。
- **分行判斷的限制**：如果段落最後一行剛好寫滿整列，會被當成自動折行而接上下一段。條列符號與日期小標題不受影響。
- **單格上限**：Google 試算表一格最多 50,000 字，超過會截斷並在執行時提示。
- **密碼外流時**：修改 `Code.gs` 的 `TOKEN`，重新部署（版本選「新版本」），並更新電腦上的環境變數。
- **Intel Mac**：新版 PyTorch 已不提供 Intel Mac 版本，docling 可能無法安裝，建議改用 Apple 晶片的 Mac 或 Windows。

---

## 七、常見問題

**`ModuleNotFoundError: No module named 'docling'`**
沒有啟用虛擬環境，或 docling 裝到另一個 Python。請先啟用虛擬環境（Mac：`source .venv/bin/activate`；Windows：`.\.venv\Scripts\Activate.ps1`），再執行 `python -m pip install docling`。

**Mac 出現 `error: externally-managed-environment`**
Homebrew 的 Python 不允許直接安裝套件，請依照「二、Mac 安裝與使用」建立虛擬環境。不要使用 `--break-system-packages`。

**`SSL 憑證驗證失敗`**
執行 `python -m pip install certifi`。在 Mac 上，使用 python.org 版 Python 的話，也可以執行「應用程式 → Python 3.12 → Install Certificates.command」。

**`網頁應用程式沒有回傳 JSON`**
- 確認部署的「誰可以存取」是「**任何人**」。
- 確認網址以 `/exec` 結尾，不是 `/dev`。
- 修改 `Code.gs` 後，確認已部署「**新版本**」。

**`Apps Script 回報錯誤：token 錯誤`**
`--token` 與 `Code.gs` 裡的 `TOKEN` 不一致。如果 `TOKEN` 有改過，記得重新部署。

**Windows 安裝 docling 時出現路徑過長的錯誤**
以系統管理員身分開啟 PowerShell，執行下面這行，重新開機後再安裝：
```powershell
New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1 -PropertyType DWORD -Force
```

**Windows 出現 huggingface symlink 警告**
不影響執行，可以忽略。

**某些儲存格的換行不對**
執行時加 `-v`，並記下案號與 PDF 原文，再調整判斷門檻。
