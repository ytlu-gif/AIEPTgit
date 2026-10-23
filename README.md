# AI 互動式英語練習工具 (AIEPT)

Windows 桌面程式，可使用 AI 產生雙語英語學習卡與語音，也提供 AI 英文對話練習。

## 執行原始碼

需要安裝 Python 3.10 或更新版本。於專案資料夾開啟 PowerShell：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

第一次使用時，在程式的「設定 → 偏好設定」設定 API Key。API Key 和對話紀錄存在本機 `config.json`、`chat_history.json`，不應提交到 Git。

## 建立 Windows 可攜式版本

在專案根目錄執行：

```powershell
python -m PyInstaller --noconfirm --clean --onedir --windowed --collect-all sounddevice --icon=app_icon.ico --name "AIEPT" app.py
```

輸出於 `dist\AIEPT\`。請完整提供該資料夾，不要只散發 `AIEPT.exe`。

## 建立安裝程式

先成功建立並測試 `dist\AIEPT\`，再以 Inno Setup 開啟 `AIEPT.iss` 並編譯。安裝檔會輸出為 `AIEPT_Setup_v2.7.exe`。

## 專案檔案

- `app.py`：目前使用中的程式原始碼。
- `AIEPT.spec`：PyInstaller 打包設定。
- `AIEPT.iss`：Inno Setup 安裝程式設定。
- `requirements.txt`：Python 套件需求。
- `pyinstaller.txt`：Windows 打包步驟。

`config.json`、聊天紀錄、生成的教材／音訊，以及 EXE 和打包目錄皆由 `.gitignore` 排除。

## GitHub 提交與發行注意事項

- 提交原始碼、圖示、README、套件清單與打包設定即可；不要強制加入被 `.gitignore` 排除的檔案。
- API Key 與聊天紀錄只保留在本機。不要將真實金鑰寫進程式、範例、截圖或問題回報。
- `.venv/`、建置目錄、安裝 EXE、生成教材與音訊均不納入原始碼版本管理。
- 可攜式版本必須從乾淨的建置輸出產生；散發前確認其中沒有 `config.json`、`chat_history.json`、`.env` 或私人教材與錄音。
- Inno Setup 設定已排除上述敏感設定及常見私人輸出；這些規則不會修改以前編譯的安裝檔。2026-10-21 已重新編譯並檢查 `AIEPT_Setup_v2.7.exe`；檢查結果見 `INSTALLER_VERIFICATION.md`。後續每次發行仍需重新檢查。
- 如金鑰曾公開或隨安裝包提供，請至服務供應商撤銷並重新產生。
- 公開專案前確認原始碼、圖示及第三方元件的使用權；本專案原始碼採用 MIT License，第三方元件仍依各自授權條款提供。

提交前可在專案根目錄檢查：

```powershell
git status --short --untracked-files=all
git status --short --ignored
git add --dry-run .
```

確認只列出預期的原始碼與文件後，再自行提交及設定遠端。公開或私人儲存庫都不應包含金鑰。

## 版權與授權

Copyright (c) 2026 Yu-Tsang Lu.

本專案原始碼採用 [MIT License](LICENSE)。使用、修改或散布本軟體時，請依授權條款保留版權聲明與授權文字。

第三方套件及隨附元件依各自授權條款提供；AI 與語音服務另依其服務條款使用。
