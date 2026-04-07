# PPT v3.0 全面視覺重建設計

## 目標

將 v2.0「太醜」的版面徹底重建為「視覺衝擊力」風格的科技發表會簡報。修正所有已知 bug（元素超出畫面、重疊、不對稱），建立集中式 layout grid，升級配色為更飽和的電光藍+螢光青綠方案。

## 核心修正

v2.0 有三類關鍵問題：

1. **元素超出畫面**：theme.py 註解誤寫為 7.5"，實際高度為 5.625"。封面日期 (y=6.6")、裝飾線 (y=6.55")、Section Break 標籤 (y=6.2") 完全不可見。Stack diagram 8 層超出底部。
2. **佈局不一致**：各 builder 使用 ad-hoc Inches() 值，左右 margin 不對稱（two_col: 0.25" vs 0.95"）。
3. **視覺風格不足**：section break 只有平面矩形，flow 箭頭幾乎不可見（0.18" gap），table 缺少背景面板。

## 設計方案：深藍科技風強化（方案 A）

### 1. 集中式 Layout Grid（theme.py）

```
SLIDE_WIDTH    = 10"      (不變)
SLIDE_HEIGHT   = 5.625"   (修正註解)

TITLE_BAR_H    = 1.0"     (標題列高度)
MARGIN_H       = 0.4"     (左右對稱 margin)
CONTENT_TOP    = 1.12"    (title bar + 0.12" gap)
CONTENT_BOTTOM = 4.99"    (footer 上方 0.05" gap)
FOOTER_H       = 0.585"   (底部資訊列)
CONTENT_WIDTH  = 9.2"     (10" - 2×0.4")
CONTENT_HEIGHT = 3.87"    (content area 可用高度)
GUTTER         = 0.35"    (欄間距)
NOTE_HEIGHT    = 0.30"    (備註文字高度)
NOTE_TOP       = 4.69"    (CONTENT_BOTTOM - NOTE_HEIGHT)
```

所有 builder 必須使用這些 grid 常數，禁止 ad-hoc Inches() 值。

### 2. 配色升級

| 角色 | Hex | 說明 |
|------|-----|------|
| BG_COLOR | `#1a1a2e` | 深藍底色（不變） |
| PRIMARY_COLOR | `#00b4ff` | 電光藍（升級） |
| ACCENT_COLOR | `#ff4757` | 鮮紅（升級） |
| ACCENT2_COLOR | `#00ffc8` | 螢光青綠（新增） |
| TEXT_COLOR | `#f0f0f0` | 微降亮度白（升級） |
| SUBTEXT_COLOR | `#8ab4f8` | 柔和副文字（升級） |
| PANEL_COLOR | `#0d1117` | GitHub Dark 面板色（升級） |
| PANEL_BORDER | `#30363d` | GitHub Dark 邊框色（升級） |

Section Colors（更飽和）：
- Section 1: `#00b4ff` 電光藍
- Section 2: `#a855f7` 鮮紫
- Section 3: `#00ffc8` 螢光青綠
- Section 4: `#fbbf24` 琥珀金

Title Bar Colors（暗色變體，用 _darken 計算或定義）。

### 3. 各 Builder 設計

#### 3.1 Cover 封面（幾何爆炸風）

裝飾元素分佈在四個角落，使用四種 section color：
- **左上角大色塊** (~2.5"×2.0")：section 1 電光藍
- **右上角大三角形** (~2.0" 寬)：section 2 紫（用 MSO_SHAPE.RIGHT_TRIANGLE, type 6）
- **右中菱形** (~1.2")：section 3 青綠（用 MSO_SHAPE.DIAMOND, type 4）
- **右下角小方塊** (~0.6")：section 4 金
- **全寬螢光青綠粗線** (4pt) 在標題上方
- **全寬鮮紅粗線** (4pt) 分隔標題與副標題
- **標題兩側 accent bar**：垂直的電光藍短條
- 主標題 48pt bold 電光藍，居中
- 副標題 24pt 青綠，居中
- 日期版號在 y=4.5" 附近（確保在畫面內）

#### 3.2 Section Break 過場

- 頂部全寬 section color 粗條（0.15" 高）
- 中央面板（8.0"×2.4"）在 y=1.3"，section-color 暗色填充
- 面板左側 0.08" 粗 accent bar（section color）
- 標題 44pt 白色 bold 居中
- 副標題 18pt section-color 居中
- "Section N" 文字在 y=4.8"（畫面內）
- 底部全寬水平線（section color）

#### 3.3 Bullets 子彈清單

- 使用 `add_title_bar` + `add_content_panel`
- 頂層 bullet 用 section-color `●` 前綴
- 子彈用 `›` 前綴
- 使用 grid 常數定位

#### 3.4 Two Column 雙欄

- 左右對稱：各 `(CONTENT_WIDTH - GUTTER) / 2 = 4.425"` 寬
- 各有獨立 content panel
- 中間分隔線用 section color
- 左右 margin 都是 0.4"（修正 v2.0 不對稱）

#### 3.5 Table 表格

- 新增 content panel 背景（v2.0 缺少）
- 表頭使用 section color
- 使用 grid 常數定位

#### 3.6 Flow 流程圖

- 箭頭間距 **0.4"**（v2.0: 0.18"，幾乎不可見）
- 5 項時：box_w=1.5", gap=0.4", total=9.1"
- 方塊使用圓角矩形 + section-color 暗色填充
- 描述整合在方塊內（下半部）
- 垂直置中在 content area

#### 3.7 Stack Diagram 堆疊圖

- box_h 從 0.55" → **0.42"**
- 垂直置中在 content area（8 層：total=3.58"，start_y=1.26"，end_y=4.84"）
- 層間向下箭頭 connector
- 邊框使用 section color

#### 3.8 Stack Diagram Annotated

- box_h 從 0.58" → **0.44"**
- start_y = CONTENT_TOP + 0.05"
- danger/warning 邊框 3pt（紅/黃）

#### 3.9 References 參考資料

- content panel 包裝
- 使用 grid 常數定位

### 4. 共用 Helper 修改

- `add_title_bar`：高度改用 TITLE_BAR_H (1.0")，section-color 填充+白色文字
- `add_content_panel`：預設使用 grid 常數
- `add_footer_bar`：使用 FOOTER_H 常數，確保在 5.625" 內
- `add_note`：使用 NOTE_TOP/NOTE_HEIGHT 常數，確保不與 footer 重疊

### 5. 安全規則

- 所有座標必須用 `int()` 包裝
- 不使用 `lxml.etree`，僅用 `OxmlElement` + `qn`
- content.py 結構不變（33 張），只改 version 字串
- create_ppt.py 輸出改為 `network-card-csi-v3.0.pptx`

### 6. 修改檔案

| 檔案 | 修改範圍 |
|------|---------|
| `src/theme.py` | 新增 grid 常數、升級配色、修正高度註解 |
| `src/builders.py` | 重寫所有 9 個 builder + 修改 4 個 helper |
| `src/content.py` | 僅更新版號 v1.1 → v3.0 |
| `src/create_ppt.py` | 輸出檔名改為 v3.0 |
| `tests/test_ppt.py` | 更新路徑 + 新增佈局驗證測試 |

### 7. 驗證方式

1. `py -m pytest tests/test_ppt.py -v` — 全部通過
2. `py src/create_ppt.py` — 成功生成 v3.0.pptx
3. Python 驗證無 float 座標：`re.findall(r'[xy]="[\d]+\.[\d]+"', xml)` 回傳空
4. Python 驗證無超出畫面的 shape：所有 shape.top + shape.height <= SLIDE_HEIGHT
5. 上傳 GitHub Release v3.0
