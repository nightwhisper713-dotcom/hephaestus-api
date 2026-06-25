# ⚙️ HEPHAESTUS API

幻翔設計 AI 行銷團隊 — 文件鍛造後端服務

## 功能

| 端點 | 功能 |
|---|---|
| `POST /generate/ppt` | 生成精美 PPT（6種投影片版型） |
| `POST /generate/excel` | 生成智慧 Excel（公式、凍結列、自動篩選） |
| `POST /generate/word` | 生成品牌 Word（頁首頁尾、樣式、表格） |
| `POST /generate/merge-excel` | **多個 Excel 合併**（上傳多檔，垂直疊加或分頁合併） |
| `POST /generate/bundle` | 同時生成 PPT + Excel + Word，打包 ZIP 下載 |

---

## 部署到 Render（免費）

### 步驟 1：上傳到 GitHub

```bash
git init
git add .
git commit -m "HEPHAESTUS API v1.0"
git remote add origin https://github.com/你的帳號/hephaestus-api.git
git push -u origin main
```

### 步驟 2：在 Render 建立服務

1. 前往 https://render.com 並登入（可用 GitHub 帳號）
2. 點 **New → Web Service**
3. 連接你的 GitHub repo
4. 設定如下：
   - **Name**: `hephaestus-api`
   - **Runtime**: `Node`
   - **Build Command**: `npm install`
   - **Start Command**: `npm start`
   - **Plan**: `Free`
5. 點 **Create Web Service**

等待部署完成（約 2-3 分鐘），你會得到一個網址：
```
https://hephaestus-api.onrender.com
```

### 步驟 3：更新 HTML 工具

把上面的網址貼到幻翔設計 AI 行銷團隊 HTML 工具的 HEPHAESTUS_API_URL 設定欄位即可。

---

## API 使用範例

### 生成 Excel（含公式）

```json
POST /generate/excel
{
  "filename": "幻翔銷售報表",
  "sheets": [{
    "name": "月份銷售",
    "headers": ["月份", "產品", "銷售額", "成本", "毛利"],
    "rows": [
      ["1月", "廣告招牌", 85000, 42000, 43000],
      ["2月", "商業空間", 120000, 65000, 55000]
    ],
    "addSummary": true
  }]
}
```

### 合併多個 Excel

```
POST /generate/merge-excel  (multipart/form-data)
files[]: 一月報表.xlsx
files[]: 二月報表.xlsx
files[]: 三月報表.xlsx
mode: stack          ← 垂直疊加（同格式）
outputName: Q1合併報表
```

或：
```
mode: sheet          ← 每個檔案變獨立工作表
```

### 同時生成三種格式（ZIP）

```json
POST /generate/bundle
{
  "zipName": "幻翔行銷方案包",
  "ppt": {
    "filename": "行銷提案",
    "slides": [
      {"type": "title", "title": "幻翔設計行銷提案", "subtitle": "2025年度策略"},
      {"type": "content", "title": "核心目標", "bullets": ["擴大品牌曝光", "開發新客群"]},
      {"type": "closing", "title": "立即合作", "brand": "幻翔商用設計"}
    ]
  },
  "excel": {
    "filename": "行銷預算表",
    "sheets": [{"name": "預算", "headers": ["項目","金額","備註"], "rows": [["社群廣告",5000,"FB/IG"]]}]
  },
  "word": {
    "filename": "行銷企劃書",
    "sections": [
      {"type": "h1", "text": "幻翔設計行銷企劃"},
      {"type": "p", "text": "本企劃針對台南在地中小企業市場..."}
    ]
  }
}
```

---

## PPT 投影片版型

| type | 說明 | 必要欄位 |
|---|---|---|
| `title` | 封面 | `title`, `subtitle` |
| `content` | 內容列表 | `title`, `bullets[]` |
| `stats` | 數字統計卡 | `title`, `stats[]` (`value`, `label`, `sub`) |
| `two_col` | 雙欄比較 | `title`, `left[]`, `right[]`, `leftLabel`, `rightLabel` |
| `table` | 資料表格 | `title`, `headers[]`, `rows[][]` |
| `closing` | 結語CTA | `title`, `subtitle`, `contact[]` |

---

## 注意事項

- Render 免費方案閒置 15 分鐘後休眠，第一次呼叫需等 10-30 秒喚醒
- 單檔上傳上限 20MB，多檔合計建議不超過 50MB
- 建議搭配 Claude API 使用：讓 Claude 分析素材 → 生成結構化 JSON → 呼叫本 API 生成文件
