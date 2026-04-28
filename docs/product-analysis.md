# 產品分析：AI 教材與講稿生成 APP

## MD 核心洞察

原始 MD 描述的是一個面向專業教育工作者的 AI 教學協作平台，而不是單純的投影片生成器。它的重點是把教育學理、課堂進度、時間控制、即時問答與版本管理整合到同一個備課與授課流程。

## 需求拆解

### 1. 引導式教材生成

教師輸入課題、年級、目標分鐘與學習目標後，系統要主動追問關鍵資訊，而非只提供單行 prompt。生成邏輯應採用逆向設計，先確認學習成果與評量，再生成教材。

對應 MVP：

- 課程訪談表單
- AI 主動追問
- Bloom 層次選擇
- Gagne 九大教學事件投影片序列
- 可編輯投影片草稿
- 單頁再生成

### 2. 進度感知講稿

教師每堂課前可以上傳教材，指定目前講到第幾頁，再輸入剩餘授課分鐘。系統需要理解前後文，生成接續講稿，並精準控制字數與講授節奏。

對應 MVP：

- 教材文字上傳與貼上
- 起始頁設定
- 目標分鐘設定
- WPM 語速模型
- 真實時間扣除：開場、核心講授、問答、緩衝
- 可編輯講稿

### 3. 即時課堂 AI 助理

原始文件強調課堂 AI 不應只是聊天框，而要能理解當前課堂脈絡，支援即時提問、事實查核、互動活動與收束評量。

對應 MVP：

- 課堂脈絡欄位自動帶入教材摘要
- 快速 prompt
- 即時問答紀錄
- 事實查核與 exit ticket 生成

### 4. 版本控制與匯出

教師需要保留不同班級、不同年份與不同難度的教材版本，並能在 AI 修改前後保持控制權。

對應 MVP：

- localStorage 版本保存
- 歷史版本還原
- 當前草稿與歷史版本比較
- Markdown 與 JSON 匯出

## MVP 邊界

這個版本是可操作的前端原型，預設使用本機規則生成內容，不需要 API key 或後端服務。第二階段新增了無套件 Node.js 後端代理：若提供 `OPENAI_API_KEY`，教材生成、進度講稿與即時助理會改走 OpenAI Responses API；若沒有 key，前端自動回到本機 fallback。

已在第二階段新增：

- `server.js` 作為本機 AI proxy，避免把 API key 放在瀏覽器端。
- `/api/ai/lesson` 生成結構化教材與追問。
- `/api/ai/script` 依教材、起始頁、分鐘與 WPM 生成講稿。
- `/api/ai/assistant` 生成課堂即時回應、查核點與下一步操作。

已在第三階段新增：

- `/api/parse-material` 可解析 TXT、MD、PPTX、DOCX、PDF。
- PPTX/DOCX 透過 OpenXML zip 結構抽取投影片、文件文字與講者備註。
- PDF 提供基礎文字抽取，作為非掃描 PDF 的先行支援。
- 前端保留解析後的頁碼/片段，講稿生成時會以起始頁、課題與學習目標做簡單檢索排序。

已在第四階段新增：

- `/api/export-pptx` 可把目前教材草稿匯出為基礎 PowerPoint。
- PPTX 內含投影片標題、教學事件、Bloom 層次、分鐘、備註摘要與 AI-assisted 標記。
- 前端新增生成紀錄（audit log），記錄 AI/本機生成、教材解析、局部修改、版本保存與匯出。
- Markdown 與 JSON 匯出會包含 AI 生成透明度或 audit log，強化 human-in-the-loop 控制。

未在本版本實作：

- 掃描 PDF OCR
- RAG 向量資料庫
- 高階 PPTX 模板、圖片、動畫與完整講者備註
- LMS 登入與成績回傳

## 建議技術路線

1. 前端維持目前工作台資訊架構，改成 React 或 Vue 後可拆成 `LessonBuilder`、`ScriptGenerator`、`LiveAssistant`、`VersionLibrary` 四個模組。
2. 後端新增 AI orchestration API，將表單狀態轉成可審核的 prompt chain。
3. 文件解析層支援 PPTX、PDF、DOCX，抽出頁碼、標題、講者備註與文字區塊。
4. RAG 層按教師帳戶隔離，避免不同租戶教材交叉檢索。
5. 匯出層加入 PPTX 與 LMS package。
