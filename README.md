# EduScript AI Studio

根據 `教材講稿AI生成APP功能.md` 分析後製作的靜態前端 MVP。這個 APP 將原始報告中的「AI 教材生成、進度感知講稿、即時課堂助理、版本控制與匯出」整理成一個可直接操作的教師工作台。

第二份 `deep-research-report.md` 已整合到產品方向：APP 不再只是教材生成器，而是朝「教材作業系統」演進，加入教材中介格式、來源追溯、發布版教材與學生端 grounded 問答。

## 功能

- 引導式課程訪談：輸入課題、科目、對象、分鐘、學習目標與班情。
- 教材生成：依 Bloom's Taxonomy 與 Gagne's Nine Events 產出可編輯投影片草稿。
- 局部再生成：選擇單頁投影片並輸入修改意見。
- 進度講稿：貼上或上傳教材文字，設定起始頁與分鐘，自動計算 WPM、核心講授時間與講稿。
- 教材解析：伺服器模式支援 TXT、MD、PPTX、DOCX、PDF 的基礎文字抽取。
- 課表解析：伺服器模式支援 XLSX 工作表文字抽取，可作為課次/時間表素材。
- 即時助理：基於目前課堂脈絡生成回答、互動問題、事實查核提示與 exit ticket。
- 學生問答：只根據教師發布版本回答，並列出來源投影片或教材片段。
- 版本保存：使用 localStorage 儲存教材版本，支援還原與比較。
- 發布治理：教師可將草稿發布為學生端可見版本。
- 角色模式：支援教師、助教、學生、管理者的前端權限模擬。
- 本機 citation index：發布時建立 chunk id、source hash、搜尋向量與 confidence。
- 治理指標：顯示發布狀態、citation chunk 數、QA 有據率、拒答率與需老師介入次數。
- 匯出：支援 Markdown、JSON 與伺服器模式 PPTX。
- AI 透明度：記錄教材生成、講稿生成、解析、匯出、版本保存等事件。

## 使用方式

### 離線原型模式

直接用瀏覽器打開：

```text
index.html
```

這個模式不需要後端或 API key，會使用本機規則生成教材與講稿。

### AI 伺服器模式

需要 Node.js 18 或以上。

1. 複製 `.env.example` 為 `.env`
2. 在 `.env` 填入 `OPENAI_API_KEY`
3. 啟動伺服器

```bash
node server.js
```

如果你的環境有 npm，也可以用 `npm start`。

開啟：

```text
http://localhost:4173
```

有 API key 時，教材生成、講稿生成與即時助理會透過 `server.js` 呼叫 OpenAI Responses API。沒有 API key 時，前端會自動回到本機規則生成。

### 上傳教材

- 離線模式：支援純文字類檔案，例如 `.txt`、`.md`、`.csv`、`.json`。
- 伺服器模式：額外支援 `.pptx`、`.docx`、`.xlsx`、`.pdf`。
- PPTX/DOCX 會從 OpenXML 結構抽取文字與講者備註。
- XLSX 會抽取 workbook / worksheet 文字，適合匯入課表、週次與教學時間安排。
- PDF 目前是基礎文字抽取；掃描圖像型 PDF 仍需要 OCR。
- 解析後會保留頁碼/片段，講稿生成時會根據起始頁與課題關鍵字選取相關片段。

### 發布與學生問答

教師完成教材與講稿後，可以按「發布教材」。學生問答頁只會使用已發布快照，不會讀取教師草稿。回答會分成：

- `教材有據`：找到足夠來源，附投影片或教材片段。
- `尚未在教材找到`：未找到足夠依據，提示學生請老師補充。
- `尚未發布`：教師還未發布教材。

學生可標記「有幫助」或「需要老師」，回饋會進入生成紀錄。

### 治理與檢索

發布教材時，系統會建立本機 citation index。每個來源片段包含：

- `id`
- `sourceHash`
- `type`
- `label`
- `preview`
- 簡易搜尋向量

學生提問時會用關鍵詞重疊與 cosine similarity 近似排序來源，並顯示 confidence。這不是正式向量資料庫，但資料結構已為未來接 embeddings / pgvector / OpenAI File Search / Azure AI Search 做準備。

### 匯出 PowerPoint

伺服器模式可匯出 `.pptx`。匯出的簡報包含：

- 每頁投影片標題
- Gagne 教學事件
- Bloom 層次
- 建議分鐘
- 教師備註摘要
- `AI-Assisted Generation / Human review required` 標記

目前 PPTX 匯出是基礎 OpenXML 產生器，適合交付與二次編輯；還未包含複雜動畫、母片設計或完整講者備註結構。

## 檔案結構

```text
.
├── index.html
├── styles.css
├── app.js
├── server.js
├── package.json
├── .env.example
├── docs/
│   └── product-analysis.md
└── README.md
```

## 下一步可擴展

- 加入真正的向量資料庫與多文件 RAG。
- 增加 OCR，支援掃描 PDF 與圖片教材。
- 增強 PPTX 匯出：主題模板、講者備註、圖片、圖表與動畫。
- 增加 LMS 整合，例如 Canvas 或 Google Classroom。
- 增加教師審核紀錄與 AI 生成透明度標記。
