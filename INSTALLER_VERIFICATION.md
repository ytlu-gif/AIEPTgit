# 安裝包驗證紀錄

日期：2026-10-21

## 舊安裝包

舊版 `AIEPT_Setup_v2.7.exe` 解包後確認含有 `config.json`，其中存在非空 API Key 設定值。不得散發。舊檔保留在已被 Git 忽略的 `build/installer-audit/AIEPT_Setup_v2.7.before-cleanup.exe`。本紀錄不包含金鑰值；如舊檔曾對外提供，請撤銷相關金鑰並換新。

## 重建與驗證

以本機 Inno Setup 7 的 ISCC，使用目前的 `AIEPT.iss` 成功重新編譯。使用 innounp 2.71.2 檢查實際安裝包：

- 安裝包清單未包含 config.json、chat_history.json、.env、常見私鑰檔、教材、錄音或日誌。
- 完整壓縮內容測試通過。
- 解包後 AIEPT.exe 的 SHA256 與 dist/AIEPT/AIEPT.exe 相符。
- 專案根目錄 Python 原始碼語法檢查通過。

重建安裝包 SHA256：

```text
C93C322E5C6E6494EABBED8B5E2715C9BFAE4766E40DE17B0C10181B5281D5E2
```

## 驗證範圍

此次驗證針對安裝包內容、打包完整性及原始碼語法。沒有執行安裝、桌面介面或付費 AI API 功能測試，也沒有執行完整惡意程式掃描。安裝包未具數位簽章；dist 主程式是否由目前原始碼建置，未重新驗證。EXE 與舊備份均不納入 Git。
