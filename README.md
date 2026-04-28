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

直接用瀏覽器打開：

```text
index.html
```

或者用任何靜態伺服器開啟本資料夾。

## 檔案結構

```text
.
├── index.html
├── styles.css
├── app.js
├── docs/
│   └── product-analysis.md
└── README.md
```

## 下一步可擴展

- 接入真正的 LLM API，將目前的本機規則生成替換成後端 AI 生成。
- 支援 PPTX/PDF 解析與向量檢索。
- 增加 PPTX 原生匯出。
- 增加 LMS 整合，例如 Canvas 或 Google Classroom。
- 增加教師審核紀錄與 AI 生成透明度標記。
