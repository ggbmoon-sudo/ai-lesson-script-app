# EduScript AI Studio

根據 `教材講稿AI生成APP功能.md` 分析後製作的靜態前端 MVP。這個 APP 將原始報告中的「AI 教材生成、進度感知講稿、即時課堂助理、版本控制與匯出」整理成一個可直接操作的教師工作台。

第二份 `deep-research-report.md` 已整合到產品方向：APP 不再只是教材生成器，而是朝「教材作業系統」演進，加入教材中介格式、來源追溯、發布版教材與學生端 grounded 問答。

## 功能

- 引導式課程訪談：輸入課題、科目、對象、分鐘、學習目標與班情。
- 全年課程包：一次生成整個學年的 Lecture/PPT、影片錄製小時、CA Lab Series 與 CA/EA 評核規劃。
- Timetable：整合 Lecture、CA Lab Series、Assessment 的週次、小時、交付物與依賴關係。
- Lab / Assessment 內容生成：每個 Lab 與評核項可生成 instructions、steps、rubric 與驗收標準。
- 教材共創：依 Bloom's Taxonomy 與 Gagne's Nine Events 產出每頁可編輯 PPT Prompt，而不是逐字講稿。
- 局部再生成：選擇單頁投影片並輸入修改意見。
- 進度講稿：把已生成 PPT / PPTX / PPT prompt 放入素材區，設定起始頁與分鐘，自動計算 WPM、核心講授時間與每堂講稿。
- 教材解析：伺服器模式支援 TXT、MD、PPTX、DOCX、PDF 的基礎文字抽取。
- 課表解析：伺服器模式支援 XLSX 工作表文字抽取，可作為課次/時間表素材。
- 即時助理：基於目前課堂脈絡生成回答、互動問題、事實查核提示與 exit ticket。
- 學生問答：只根據教師發布版本回答，並列出來源投影片或教材片段。
- 版本保存：使用 localStorage 儲存教材版本，支援還原與比較。
- 發布治理：教師可將草稿發布為學生端可見版本。
- 角色模式：支援教師、助教、學生、管理者的前端權限模擬。
- 本機 citation index：發布時建立 chunk id、source hash、搜尋向量與 confidence。
- 治理指標：顯示發布狀態、citation chunk 數、QA 有據率、拒答率與需老師介入次數。
- 個人備份：支援本機 JSON 匯入/匯出、Google Drive 雲端備份與還原、未備份提醒與自動備份開關。
- 匯出：支援 Markdown、JSON、伺服器模式 PPTX，以及一鍵 Course Pack ZIP。
- AI 狀態燈：顯示 Gemini / OpenAI / 本機 fallback，並可在 APP 內重新檢查連線。
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

如果你想讓 Google Drive OAuth Client ID 自動出現在 APP，可在 `.env` 加入：

```text
GOOGLE_DRIVE_CLIENT_ID=你的_client_id.apps.googleusercontent.com
```

Gemini 也可作為 AI 後端。如果你沒有 OpenAI token，在 `.env` 這樣設定：

```text
AI_PROVIDER=gemini
GEMINI_API_KEY=你的 Gemini API key
GEMINI_MODEL=gemini-3-pro-preview
GEMINI_THINKING_LEVEL=high
GEMINI_TEMPERATURE=0.25
GEMINI_SCRIPT_MAX_OUTPUT_TOKENS=32768
```

`AI_PROVIDER=auto` 會先用 OpenAI key；沒有 OpenAI key 時會改用 Gemini key。兩者都沒有時，APP 仍會使用本機規則生成。高品質完整講稿建議使用 `gemini-3-pro-preview` + `GEMINI_THINKING_LEVEL=high`；如果想節省延遲或成本，可改為 `gemini-3-flash-preview`。Gemini 2.5 系列可用 `GEMINI_THINKING_BUDGET=-1` 啟用 dynamic thinking。

講稿生成流程已改為兩階段：先把 PPTX 解析成乾淨的 `slide_json`，每頁保留 `slide_no`、`slide_title`、`slide_subtitle`、`slide_body`、`visual_description`、`speaker_notes`、`source_type` 與 `extracted_from`；再把完整 `slide_json`、課程資訊、訪談資料、學習目標、對象、時長與風格交給 Gemini 生成逐頁教師口語講稿。APP 會避免把 PPT Prompt、compiler prompt、debug log 或版本紀錄混入正式講稿。

AI 狀態燈會顯示 Gemini 是否屬於高階模型、目前 temperature，以及講稿專用 max output tokens。若使用 Flash / Lite 類模型，APP 會提示完整技術講稿較建議改用 Pro / Thinking 類模型。

### 部署版

本 repo 已加入 Render 部署設定：

- `render.yaml`
- `.node-version`
- `docs/deployment.md`

Render 會以 Node web service 執行 `server.js`，保留 AI proxy、PPTX 匯出、教材解析與 Google Drive 設定載入功能。部署後要在 Render 設定環境變數，不要把真實 key 提交到 GitHub：

```text
AI_PROVIDER=auto
OPENAI_API_KEY=你的 OpenAI key，可留空
OPENAI_MODEL=gpt-5.2
GEMINI_API_KEY=你的 Gemini key，可留空
GEMINI_MODEL=gemini-3-pro-preview
GEMINI_THINKING_LEVEL=high
GOOGLE_DRIVE_CLIENT_ID=你的 Google OAuth Client ID
PUBLIC_BASE_URL=https://你的服務.onrender.com
```

部署後，記得到 Google OAuth Client 的 Authorized JavaScript origins 加入 Render 網址，例如：

```text
https://你的服務.onrender.com
```

詳細步驟見 `docs/deployment.md`。

### 上傳教材

- 離線模式：支援純文字類檔案，例如 `.txt`、`.md`、`.csv`、`.json`。
- 伺服器模式：額外支援 `.pptx`、`.docx`、`.xlsx`、`.pdf`。
- PPTX/DOCX 會從 OpenXML 結構抽取文字與講者備註。
- XLSX 會抽取 workbook / worksheet 文字，適合匯入課表、週次與教學時間安排。
- PDF 目前是基礎文字抽取；掃描圖像型 PDF 仍需要 OCR。
- 解析後會保留頁碼/片段，講稿生成時會根據起始頁與課題關鍵字選取相關片段。

### 全年課程包

「年度規劃」頁可輸入整個學年的總目標，例如進階 Linux、CKA/CKAD、EKS、Rancher、CA Lab 與 Skill Test。系統會生成：

- Lecture / PPT 清單：預設 13 小時影片，拆成逐 deck 錄影與 PPT 製作任務。
- CKA / CKAD 去重策略：共用架構與 kubectl 基礎只保留一次，管理員與開發者視角分開。
- CA Lab Series：Ubuntu VM、Ansible、Minikube CKA、Minikube CKAD、AWS Academy EKS、Isakei / Rancher。
- Assessment：CA 筆試與 Lab checkpoint、EA Isakei 作業、no-hint Skill Test 與 public endpoint 規則。

每個 Lecture deck 可按「送到 PPT 流程」，轉入「教材共創」生成可編輯投影片，再用伺服器模式匯出 PPTX。

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

### 個人備份與 Google Drive

本機 JSON 匯出/匯入可用於手動備份。若要使用 Google Drive：

1. 到 Google Cloud Console 建立專案並啟用 Google Drive API。
2. 建立 OAuth 2.0 Web application client。
3. 在 Authorized JavaScript origins 加入 `http://localhost:4173`。
4. 用 `node server.js` 啟動 APP，開啟 `http://localhost:4173`。
5. 在「版本匯出」頁貼上 OAuth Client ID，按「連接 Drive」。

APP 使用 Google Identity Services 取得使用者授權，並以 Drive API `drive.file` scope 上傳 JSON 備份。這個 scope 只讓 APP 存取它自己建立或使用者授權開啟的檔案，不會讀取整個雲端硬碟。前端只需要 OAuth Client ID，請不要把 Google Client Secret 或任何 API key 貼入 APP 或提交到 GitHub。

「儲存 / 發布後自動備份」開啟後，只要 Drive token 仍有效，儲存版本、發布教材、生成講稿等重要操作會自動排程備份；若 token 過期，系統會保留「待備份」狀態，重新按「連接 Drive」後即可再備份。

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
├── render.yaml
├── .env.example
├── .node-version
├── docs/
│   ├── deployment.md
│   ├── deep-report-integration.md
│   └── product-analysis.md
└── README.md
```

## 下一步可擴展

- 加入真正的向量資料庫與多文件 RAG。
- 增加 OCR，支援掃描 PDF 與圖片教材。
- 增強 PPTX 匯出：主題模板、講者備註、圖片、圖表與動畫。
- 增加 LMS 整合，例如 Canvas、Google Classroom 或 Moodle。
- 增加教師審核紀錄與 AI 生成透明度標記。

## Gamma PPT Connector

The app now includes a Gamma PPT connector for future direct deck generation. The frontend sends the current per-slide PPT prompts to the local `server.js` proxy, and the API key stays in environment variables only.

```text
GAMMA_API_KEY=your_gamma_api_key_here
GAMMA_EXPORT_AS=pptx
GAMMA_TEXT_MODE=generate
GAMMA_THEME_ID=
GAMMA_FOLDER_IDS=
```

If `GAMMA_API_KEY` is empty, the `Gamma PPT` button downloads a Gamma-ready Markdown prompt instead. You can paste that prompt into Gamma manually now, then add a Gamma Pro/Ultra/Teams/Business API key later to generate through `POST /api/gamma/generate`.
