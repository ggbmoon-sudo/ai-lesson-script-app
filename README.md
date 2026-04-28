# EduScript AI Studio

根據 `教材講稿AI生成APP功能.md` 分析後製作的靜態前端 MVP。這個 APP 將原始報告中的「AI 教材生成、進度感知講稿、即時課堂助理、版本控制與匯出」整理成一個可直接操作的教師工作台。

## 功能

- 引導式課程訪談：輸入課題、科目、對象、分鐘、學習目標與班情。
- 教材生成：依 Bloom's Taxonomy 與 Gagne's Nine Events 產出可編輯投影片草稿。
- 局部再生成：選擇單頁投影片並輸入修改意見。
- 進度講稿：貼上或上傳教材文字，設定起始頁與分鐘，自動計算 WPM、核心講授時間與講稿。
- 即時助理：基於目前課堂脈絡生成回答、互動問題、事實查核提示與 exit ticket。
- 版本保存：使用 localStorage 儲存教材版本，支援還原與比較。
- 匯出：支援 Markdown 與 JSON。

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

- 支援 PPTX/PDF 解析與向量檢索。
- 增加 PPTX 原生匯出。
- 增加 LMS 整合，例如 Canvas 或 Google Classroom。
- 增加教師審核紀錄與 AI 生成透明度標記。
