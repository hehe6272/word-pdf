# 軟體介紹

## 📜 簡介

這款 **Word 轉 PDF 小工具** 是一個基於 Python 和 Tkinter 開發的應用程式，旨在幫助用戶快速將 Microsoft Word 文檔 (.docx) 轉換成 PDF 格式。它提供簡單的圖形界面，無需繁瑣的設置和操作，只需幾個步驟即可完成文檔轉換。此工具適用於需要批量處理文件的用戶，尤其是想要簡化日常轉檔流程的使用者。

## 🛠️ 技術架構

- **Python**：主要程式語言，確保跨平台兼容。
- **Tkinter**：用於開發簡單直觀的圖形界面（GUI）。
- **pywin32**：用於與 Microsoft Word 應用程式互動，實現文檔轉換的核心功能。

## 🚀 使用方法

1. **啟動程式**：執行下載的程式文件，進入 GUI 界面。
2. **選擇 Word 文件**：點擊“選擇 Word 檔案”按鈕，選擇您想要轉換的 Word 文檔。
3. **選擇儲存位置**：選擇 PDF 檔案的儲存位置。程式會自動為轉換後的 PDF 文件命名為與原 Word 檔案相同的名稱。
4. **開始轉換**：點擊“開始轉換”按鈕，程式將自動進行轉換，並顯示進度條，讓您了解轉換的進度。
5. **完成轉換**：當轉換完成後，您可以在選定的儲存位置找到轉換後的 PDF 檔案。

## 🖥️ 技術細節

- 程式通過 `pywin32` 庫與 Word 進行互動，將 Word 檔案的內容轉換為 PDF 格式。這樣的做法不僅保持了格式的一致性，還能有效避免轉換過程中的格式錯亂。
- 使用 `Tkinter` 打造的簡單界面讓操作變得非常直觀，即便是技術小白也能快速上手。
- 程式的錯誤處理功能能在轉換過程中提供反饋，確保用戶能夠清楚了解轉換過程中可能出現的問題。

## 🎨 預覽界面

下圖展示了軟體的 GUI 界面，簡單而直觀：

![Word to PDF Tool GUI]([https://via.placeholder.com/600x400](https://private-user-images.githubusercontent.com/98233537/416082134-261eeda6-5890-42dc-9023-ec8f3de6b665.png?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTUiLCJleHAiOjE3NDAzNzg2MDksIm5iZiI6MTc0MDM3ODMwOSwicGF0aCI6Ii85ODIzMzUzNy80MTYwODIxMzQtMjYxZWVkYTYtNTg5MC00MmRjLTkwMjMtZWM4ZjNkZTZiNjY1LnBuZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFWQ09EWUxTQTUzUFFLNFpBJTJGMjAyNTAyMjQlMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjUwMjI0VDA2MjUwOVomWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPTVkMjlkZDNjYjI0YjcyMmQ5MmUyNjc4OGZlNjMwMTI3NGM2ZjJjMmQ5ODQ4NjZjZjNhYjE1Mjc2YTg0ZDQyNDkmWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0In0.yCCx106-GVy3j2ggP7yo8WbdVHMTI_ISuqxGEgV8r6w))

## 📅 更新日志

- **v1.0**: 初始版本，基本的 Word 轉 PDF 功能。
- **v1.1**: 增加了錯誤處理功能，改善了進度條顯示。

## 📝 使用者反饋

我們非常重視使用者的反饋，若您在使用過程中有任何建議或問題，請隨時在 GitHub 上創建問題（Issue）或聯繫我們。

