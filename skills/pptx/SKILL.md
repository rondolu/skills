---
name: pptx
description: "Use this skill any time a .pptx file is involved in any way — as input, output, or both. This includes: creating slide decks, pitch decks, or presentations; reading, parsing, or extracting text from any .pptx file (even if the extracted content will be used elsewhere, like in an email or summary); editing, modifying, or updating existing presentations; combining or splitting slide files; working with templates, layouts, speaker notes, or comments. Trigger whenever the user mentions \"deck,\" \"slides,\" \"presentation,\" or references a .pptx filename, regardless of what they plan to do with the content afterward. If a .pptx file needs to be opened, created, or touched, use this skill."
license: Proprietary. LICENSE.txt has complete terms
---
---
# PPTX 技能

## 快速參考

| 任務 | 指引 |
|------|-------|
| 讀取／分析內容 | `python -m markitdown presentation.pptx` |
| 從範本編輯或建立 | 閱讀 [editing.md](editing.md) |
| 從頭建立 | 閱讀 [pptxgenjs.md](pptxgenjs.md) |
| 檢視／切換設計範本 | [fetch_template_from_ppt.yaml](fetch_template_from_ppt.yaml)（預設）· [fetch_template_with_backgrounds.yaml](fetch_template_with_backgrounds.yaml)（背景版）· [cathay_template.yaml](cathay_template.yaml)（備用） |

---

## 輸出 Bundle 規則（必須遵守）

**每一次產生簡報前，先建立一個與簡報同名的資料夾，並把所有產物都集中在該資料夾內。**

必須使用以下結構：

```text
DeckName/
	DeckName.pptx
	assets/
		...所有外部圖片、圖示、SVG、匯入素材...
	assets/resource-manifest.json
```

必要規則：
1. 資料夾名稱必須與簡報檔名相同（不含 `.pptx` 副檔名）。
2. 最終簡報檔必須放在該資料夾根目錄，檔名為 `DeckName.pptx`。
3. 所有用到的外部資源都必須一起保存到 `assets/`，包含圖片、icon、SVG、裁切後圖檔、暫存轉檔結果。
4. 若 icon 是程式動態產生，必須先輸出成實體檔放進 `assets/`，再嵌入簡報；不要只保留記憶體中的 base64。
5. 驗證、轉圖、縮圖等衍生輸出，也應優先寫入這個資料夾內，避免產物散落工作區。

---

## 讀取內容

```bash
# 文字擷取
python -m markitdown presentation.pptx

# 視覺概覽
python scripts/thumbnail.py presentation.pptx

# 原始 XML
python scripts/office/unpack.py presentation.pptx unpacked/
```

---

## 編輯工作流程

**完整細節請閱讀 [editing.md](editing.md)。**

1. 使用 `thumbnail.py` 分析範本
2. 解包 → 操作投影片 → 編輯內容 → 清理 → 重新打包

---

## 從頭建立

**完整細節請閱讀 [pptxgenjs.md](pptxgenjs.md)。**

當沒有可用的範本或參考簡報時使用此方式。

---

## 簡報規劃（建議先做）

如果使用者已經提供完整大綱、頁數與素材，可以直接進入設計與產生流程。

如果使用者只給了主題、幾個關鍵字，或只說「幫我做一份簡報」，請先完成最小必要的規劃，再開始寫投影片。

### 需求澄清清單

在動手前，至少確認以下幾件事：

| 問題 | 為什麼要問 |
|------|------------|
| 這份簡報的受眾是誰？使用情境是提案、內部報告、分享還是教學？ | 決定語氣、深度與內容密度 |
| 預計講多久，或希望控制在幾頁內？ | 幫助規劃頁數與每頁資訊量 |
| 是否已有原始素材，例如文件、舊簡報、數據、連結或圖表？ | 決定是整理素材還是從零搭建內容 |
| 是否有圖片、截圖或必須保留的視覺資產？ | 影響版型選擇與內容分配 |
| 是否有硬性限制，例如必須包含的數字、不能省略的章節、不能出現的元素？ | 避免後期大幅返工 |

若以上資訊不完整，請先補最影響結構的 1-3 項，再開始規劃。

### 大綱協助（使用者沒有大綱時）

當使用者沒有現成大綱時，先用敘事弧搭出骨架，再把素材填進去：

```text
Hook（鉤子）        → 1 頁   : 用反差、問題或關鍵數字把人拉進來
Context（定調）     → 1-2 頁 : 說明背景、問題脈絡、為什麼現在要講
Core（主體）        → 3-5 頁 : 放核心內容、方法、案例、流程或比較
Shift（轉折）       → 1 頁   : 提出新觀點、關鍵變化或決策點
Takeaway（收束）    → 1-2 頁 : 總結、行動建議、下一步或結論
```

實務上可先決定以下三件事，再開始寫投影片：

1. 故事順序：內容應該先講問題、方法、案例，還是先講結論再回推。
2. 頁數配置：把有限頁數優先留給最重要的主體內容，不要平均分配。
3. 內容形態：每一頁主要是標題、比較、流程、結構、圖像，還是結尾收束。

若要把敘事弧直接對應到現有範本，通常可用以下方式起稿：

| 敘事階段 | 優先考慮的範本 |
|----------|----------------|
| Hook | `Cover / Title Slide`、`Agenda / Table of Contents` |
| Context | `Agenda / Table of Contents`、`Content - Left Panel + Right Multi-Section` |
| Core | `Content - Horizontal Row`、`Content - Flow Diagram`、`Content - Comparison Table`、`Content - Three-Card Comparison` |
| Shift | `Content - Section Divider`、`Content - Comparison Table` |
| Takeaway | `Closing / Thank You`、`Content - Section Divider` |

### 內容先行，不要先堆版型

在選版型前，先判斷這頁的任務是什麼：

1. 要建立情境還是下結論。
2. 要解釋步驟還是比較選項。
3. 要放真實數據還是用圖像建立記憶點。

同一份 deck 不需要每頁都不同，但也不要連續多頁用同一種結構硬塞不同類型的內容。

---

## 設計系統（必須遵守）

**在產生任何簡報之前，您必須讀取當前設計範本 YAML 並套用其規格。**
設計範本完整定義了顏色、字型、版面配置及每張投影片的元素規格——請勿使用 YAML 中未定義的值。

**預設範本**（在撰寫任何投影片程式碼前，請先讀取此檔案）：
```
.github/skills/pptx/fetch_template_from_ppt.yaml
```

**背景版範本**（封面/內容頁使用不同背景圖時）：
```
.github/skills/pptx/fetch_template_with_backgrounds.yaml
```

YAML 結構如下所示：

| YAML 區段 | 定義內容 |
|--------------|-----------------|
| `design_system.color_palette` | 所有允許的十六進位色碼及使用規則——**不允許使用其他顏色** |
| `design_system.typography` | 字型家族、大小層級、字重（全部為粗體）及各文字角色的顏色 |
| `design_system.layout_rules` | 投影片尺寸（960×540pt = 10"×5.625"）、邊距、格線及 UI 元素樣式 |
| `layout_rules.recurring_elements` | 依模板決定的裝飾元素（背景版模板可為空陣列） |
| `slide_templates` | 具有精確元素位置與樣式的具名版面配置模式 |
| `instructions_for_generation` | 必要規則：語調、禁止元素、色彩規範、間距 |

### 必要閱讀工作流程

在撰寫任何投影片程式碼或 XML 之前：
1. **選擇並讀取**對應範本（標準版：`.github/skills/pptx/fetch_template_from_ppt.yaml`；背景版：`.github/skills/pptx/fetch_template_with_backgrounds.yaml`）
2. **擷取**相關設計值（顏色、字型、版面規格）
3. **套用**這些值於 PptxGenJS 呼叫或 XML 編輯——YAML 到程式碼的對應指南請參閱 [pptxgenjs.md](pptxgenjs.md)

### 裝飾元素策略（依模板而定）

定義於 `layout_rules.recurring_elements`。是否套用固定裝飾，取決於你選的模板：

| 模板 | 裝飾策略 |
|---------|----------|
| `fetch_template_from_ppt.yaml` | 保留 recurring decorators |
| `fetch_template_with_backgrounds.yaml` | 不使用固定頂部/底部/左側裝飾線 |

標準版 recurring decorators 參考如下：

| 元素 | 位置 | 樣式 |
|---------|----------|-------|
| 頂部細長條 | x=0, y=0, 全寬, h≈5pt | 深海軍藍填充（`0D1F33`） |
| 底部三段式條 | y≈518pt, h≈22pt, 三等分 | 由左至右漸層為漸淺的深藍色 |
| 左側垂直強調線 | x=0, y≈5pt, h≈513pt, w=5–8pt | 電光藍 `186AFF` |
| 頁面／章節編號標籤 | 左上角，約 80×70pt 文字框 | 40pt 粗體 `186AFF`，兩位數編碼（01/02/…） |

### 投影片範本選擇

YAML 在 `slide_templates` 下定義了具名版面配置類型。請將每張投影片的內容對應到最接近的類型：

| 範本類型 | 最適合用於 |
|---------------|---------------|
| Cover / Title Slide（封面／標題投影片） | 開場／標題頁 |
| Agenda / Table of Contents（議程／目錄） | 章節概覽，附大型 KPI 數字 |
| Content - Horizontal Row（內容－水平列） | 2–3 個並列主題；每列含圖示與正文 |
| Content - Left Panel + Right Multi-Section（內容－左側面板＋右側多節） | 深度主題，含側邊欄背景資訊 |
| Content - Section Divider（內容－章節分隔） | 章節轉換／視覺換頁 |
| Content - Comparison Table（內容－對比表格） | 錯誤 vs 正確／之前 vs 之後 |
| Content - Flow Diagram（內容－流程圖） | 循序步驟或決策樹 |
| Content - Three-Card Comparison（內容－三卡片對比） | 並排規則或選項卡片 |
| Closing / Thank You（結尾／感謝） | 摘要／結語 |

閱讀 YAML 的 `slide_templates[].elements`，可取得每個範本精確的 x/y 位置、尺寸、字型及顏色。

### 依用途挑選版型

不要只看範本名稱。選版型時，先判斷內容屬於哪一種用途，再對應到最接近的模板。

| 用途 | 優先考慮的範本 | 選型提示 |
|------|----------------|----------|
| 開場、建立主題、拋出核心問題 | `Cover / Title Slide` | 適合一個主標題加一句副標或一句主張，不要塞過多正文 |
| 章節概覽、議程、先給全局地圖 | `Agenda / Table of Contents` | 適合建立閱讀路線，也適合放章節級 KPI 或重點摘要 |
| 2-3 個平行主題並列說明 | `Content - Horizontal Row` | 適合並列概念、能力、模組、原則；不適合有明確先後順序的流程 |
| 深度說明單一主題，左側放背景、右側拆多段展開 | `Content - Left Panel + Right Multi-Section` | 適合定義、脈絡、案例拆解、問題分析 |
| 章節轉場、觀點切換、關鍵句強調 | `Content - Section Divider` | 適合拿來做呼吸頁、轉折頁，不要塞成資訊頁 |
| 錯誤 vs 正確、Before vs After、方案比較 | `Content - Comparison Table` | 有兩組以上明確對照維度時優先使用 |
| 有步驟、流程、決策樹、階段推進 | `Content - Flow Diagram` | 適合時間序列、操作流程、決策路徑；不要拿來裝平行概念 |
| 三個選項、三條規則、三種策略並列 | `Content - Three-Card Comparison` | 三張卡片要彼此等權；若某一張特別重要，應考慮換模板 |
| 收尾、總結、行動建議、致謝 | `Closing / Thank You` | 適合一句總結、3 個重點或下一步，不要再引入新主題 |

### 高資訊密度內容的選型方式

如果你的內容偏向資訊密度高、數據多、比較強，請用更嚴格的方式挑版型，但要落到目前 YAML 已存在的模板類型上：

| 內容形態 | 建議優先版型 | 原則 |
|----------|--------------|------|
| 演化比較、前後變化、版本差異 | `Content - Comparison Table`、`Content - Flow Diagram` | 先判斷重點是「差異」還是「順序」；差異用比較，順序用流程 |
| 排名、KPI、規格、數據摘要 | `Agenda / Table of Contents`、`Content - Comparison Table` | 必須有真實數據支撐；不要用純文案假裝數據頁 |
| 三層架構、模組關係、系統分工 | `Content - Left Panel + Right Multi-Section`、`Content - Horizontal Row` | 若有主從層級，用左側背景 + 右側拆段；若各模組等權，用水平列 |
| 方法論、操作步驟、決策流程 | `Content - Flow Diagram` | 流程型內容優先用流程圖，不要把步驟硬塞進卡片或表格 |
| 原則、選項、方案並列 | `Content - Three-Card Comparison`、`Content - Horizontal Row` | 三個等權選項用三卡；超過三項或每項內容較長時改用水平列 |
| 單一主張搭配證據、案例拆解 | `Content - Left Panel + Right Multi-Section` | 用左側固定觀點，右側承接證據、案例、數據或結論 |

### 版型選用原則

1. 不要為了塞內容而硬套版型；若內容形態不合，應先改寫內容或換模板。
2. 數據型頁面必須有真正可比較的數據、指標或維度，不要用長段落冒充。
3. 流程型頁面必須有順序關係；沒有順序的內容不要假裝是流程。
4. 比較型頁面必須有對照軸；若只是列點，改用水平列或多段內容頁。
5. 同一份簡報應保留節奏變化：開場、展開、比較、流程、收束不要全部長得一樣。
6. 若某頁同時想放流程、比較、案例與大量文字，通常代表這頁應拆成兩頁以上。

### 替代範本

背景圖導向模板：
```
.github/skills/pptx/fetch_template_with_backgrounds.yaml
```
若要讓封面與內容頁使用不同背景圖，請使用此模板。

原始國泰企業設計系統的備份保存於：
```
.github/skills/pptx/cathay_template.yaml
```
若要切換主題，請改讀取該 YAML 並依照相同工作流程套用其值。

---

## 品質驗證（必要步驟）

**預設有問題存在。您的任務是找出它們。**

第一次的輸出幾乎從不正確。請以找 Bug 的心態進行品質驗證，而非確認步驟。如果第一次檢查就發現零問題，代表您看得不夠仔細。

### 內容品質驗證

```bash
python -m markitdown DeckName/DeckName.pptx
```

檢查是否有遺漏內容、錯字、順序錯誤。

**使用範本時，請檢查是否有殘留的預留位置文字：**

```bash
python -m markitdown DeckName/DeckName.pptx | grep -iE "xxxx|lorem|ipsum|this.*(page|slide).*layout"
```

若 grep 有回傳結果，請在宣告完成前修正。

### 視覺品質驗證

**⚠️ 請使用子代理（SUBAGENTS）**——即使只有 2-3 張投影片也一樣。你已盯著程式碼太久，只會看見你期望看到的，而非實際存在的內容。子代理擁有全新的視角。

將投影片轉換為圖片（參見下方「轉換為圖片」章節），然後使用以下提示詞：

```
請視覺化檢查這些投影片。預設有問題存在——請找出它們。

需檢查的項目：
- 元素重疊（文字穿越圖形、線條穿越文字、堆疊元素）
- 文字溢出或在邊緣／框線處被截斷
- 裝飾線設計給單行文字，但標題換行為兩行
- 來源引用或頁腳與上方內容碰撞
- 元素間距過近（< 0.3" 間距）或卡片／節區幾乎相接
- 間距不均勻（一處有大片空白，另一處卻很擁擠）
- 距投影片邊緣邊距不足（< 0.5"）
- 欄或類似元素未對齊
- 低對比文字（例如淺灰色文字在米白色背景上）
- 低對比圖示（例如深色圖示在深色背景上，沒有對比圓圈）
- 文字框過窄導致過度換行
- 殘留的預留位置內容

針對每張投影片，列出問題或需注意之處，即使是細節也要列出。

請讀取並分析這些圖片：
1. /path/to/slide-01.jpg（預期內容：[簡短描述]）
2. /path/to/slide-02.jpg（預期內容：[簡短描述]）

回報所有發現的問題，包括細節問題。
```

### 驗證循環

1. 產生投影片 → 轉換為圖片 → 檢查
2. **列出發現的問題**（若未發現問題，請再更嚴格地重新檢查）
3. 修正問題
4. **重新驗證受影響的投影片**——修正一個問題往往會產生另一個問題
5. 重複執行，直到完整檢查一遍後不再出現新問題

**在完成至少一次修正後驗證循環之前，請勿宣告完成。**

---

## 轉換為圖片

將簡報轉換為個別投影片圖片以進行視覺檢查：

```bash
python scripts/office/soffice.py --headless --convert-to pdf DeckName/DeckName.pptx
pdftoppm -jpeg -r 150 DeckName/DeckName.pdf DeckName/slide
```

這會建立 `slide-01.jpg`、`slide-02.jpg` 等檔案。

修正後重新輸出特定投影片：

```bash
pdftoppm -jpeg -r 150 -f N -l N DeckName/DeckName.pdf DeckName/slide-fixed
```

---

## Dependencies

- `pip install "markitdown[pptx]"` - text extraction
- `pip install Pillow` - thumbnail grids
- `npm install -g pptxgenjs` - creating from scratch and PDF-to-PPTX conversion
- `pip install rapidocr_onnxruntime` - OCR for image-based PDFs (recommended, no system deps)
- `pip install pytesseract` - fallback OCR for image-based PDFs (requires Tesseract binary)
- LibreOffice (`soffice`) - PDF conversion (auto-configured for sandboxed environments via `scripts/office/soffice.py`)
- Poppler (`pdftoppm`, `pdfimages`) - PDF to images and embedded image extraction
