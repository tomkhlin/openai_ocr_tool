# 專案環境設定

## 需求

1. Python 3.10 以上
2. uv 套件管理工具

## 安裝步驟

1. 下載並安裝 [Python](https://www.python.org/downloads/)
   Windows 使用者請確保在安裝過程中勾選 "Add Python to PATH" 選項。
   Linux 和 macOS 使用者可使用系統的套件管理工具進行安裝。
2. 確認 uv 已安裝，執行以下指令：
   ```
   uv --version
   ```
    若未安裝，請參考 [uv 官方文件](https://uv.dev/) 進行安裝。
3. 使用 uv 安裝專案依賴：
   ```
   uv install
   ```
4. 執行：
   ```
   uv run main.py
   ```