# 網卡與 OS 層關係 PPT 實作計畫

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 使用 python-pptx 生成 25 張投影片的 .pptx 檔案，主題為網卡與 OS 層關係及為什麼 Windows 無法取得 CSI

**Architecture:** 單一 Python 腳本（create_ppt.py）搭配 theme.py（色彩/字型常數）與 content.py（投影片文字內容），由 main 腳本依序建立每張投影片並輸出 .pptx。

**Tech Stack:** Python 3, python-pptx, pip

---

## 檔案結構

```
network-ppt-project/
├── src/
│   ├── theme.py          # 色彩、字型、尺寸常數
│   ├── content.py        # 所有投影片文字內容（資料與程式碼分離）
│   ├── builders.py       # 投影片建構函式（每種版型一個函式）
│   └── create_ppt.py     # 主程式，依序呼叫 builders
├── output/
│   └── network-card-csi.pptx   # 最終輸出
└── tests/
    └── test_ppt.py       # 驗證投影片數量、標題、表格結構
```

---

## Task 1：安裝 python-pptx

**Files:**
- 無需建立檔案

- [ ] **Step 1: 安裝套件**

```bash
pip install python-pptx
```

Expected: `Successfully installed python-pptx-...`

- [ ] **Step 2: 驗證安裝**

```bash
python -c "from pptx import Presentation; print('OK')"
```

Expected: `OK`

---

## Task 2：建立 theme.py

**Files:**
- Create: `src/theme.py`

- [ ] **Step 1: 建立主題常數檔**

```python
# src/theme.py
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor

# 色彩
BG_COLOR       = RGBColor(0x1a, 0x1a, 0x2e)   # 深藍背景
PRIMARY_COLOR  = RGBColor(0x4a, 0x9e, 0xff)   # 主色（亮藍）
ACCENT_COLOR   = RGBColor(0xff, 0x6b, 0x6b)   # 強調（紅）
TEXT_COLOR     = RGBColor(0xff, 0xff, 0xff)   # 白色文字
SUBTEXT_COLOR  = RGBColor(0xaa, 0xcc, 0xff)   # 淡藍副文字
TABLE_HDR_BG   = RGBColor(0x0d, 0x47, 0xa1)   # 表頭深藍
TABLE_ROW_ALT  = RGBColor(0x1e, 0x2a, 0x4a)   # 交替列背景

# 字型大小
TITLE_SIZE     = Pt(36)
SUBTITLE_SIZE  = Pt(24)
BODY_SIZE      = Pt(18)
SMALL_SIZE     = Pt(14)
CODE_SIZE      = Pt(14)

# 字型名稱
FONT_TITLE     = "Calibri"
FONT_BODY      = "Calibri"
FONT_CODE      = "Consolas"

# 投影片尺寸（16:9）
SLIDE_WIDTH    = Emu(9144000)   # 10 inches
SLIDE_HEIGHT   = Emu(5143500)   # 7.5 inches
```

- [ ] **Step 2: 驗證語法**

```bash
python -c "import sys; sys.path.insert(0,'src'); import theme; print('theme OK')"
```

Expected: `theme OK`

---

## Task 3：建立 content.py

**Files:**
- Create: `src/content.py`

- [ ] **Step 1: 建立投影片內容資料**

```python
# src/content.py

SLIDES = [
    # Slide 1: 封面
    {
        "type": "cover",
        "title": "網卡與 OS 層的關係",
        "subtitle": "從 PHY/MAC 到 Windows 驅動堆疊\n——為什麼無法從 Windows 取得 CSI？",
        "author": "",
        "date": "2026-04-07",
    },
    # Slide 2: 大綱
    {
        "type": "bullets",
        "title": "大綱",
        "bullets": [
            "第一段：基礎架構（Slides 3–6）",
            "  · 802.11 RF 信號路徑",
            "  · PHY 層與 CSI 定義",
            "  · MAC 層職責",
            "  · OSI 7 層 vs Windows 元件對應",
            "第二段：Windows 驅動堆疊（Slides 7–14）",
            "  · OEM.sys / nwifi.sys / NDIS / WDI",
            "  · CSI 消失點分析",
            "第三段：替代方案（Slides 15–22）",
            "  · Linux 生態 vs Windows 限制",
            "  · 硬體選型建議",
            "第四段：結語與建議（Slides 23–25）",
        ],
    },
    # Slide 3: 802.11 RF 信號路徑
    {
        "type": "flow",
        "title": "802.11 RF 信號路徑",
        "flow_items": [
            ("天線\n(Antenna)", "接收 RF 信號"),
            ("ADC\n類比→數位", "信號數位化"),
            ("OFDM 解調\n(Demodulation)", "子載波分離"),
            ("802.11 Frame\n解碼", "封包還原"),
            ("MAC 層\n處理", "上送至驅動"),
        ],
        "note": "CSI 在 OFDM 解調後即可計算：每個子載波的複數增益（振幅 + 相位）",
    },
    # Slide 4: PHY 層 - 什麼是 CSI
    {
        "type": "two_col",
        "title": "PHY 層：什麼是 CSI？",
        "left_title": "RSSI（傳統方式）",
        "left_bullets": [
            "整體接收功率（單一數值）",
            "單位：dBm",
            "無法區分多路徑效應",
            "精度低，只反映總體信號強度",
        ],
        "right_title": "CSI（Channel State Information）",
        "right_bullets": [
            "每個 OFDM 子載波的複數增益",
            "包含振幅（Amplitude）與相位（Phase）",
            "802.11n：30 個子載波群組（20MHz）",
            "可重建多路徑傳播特性",
            "用途：室內定位、手勢辨識、呼吸偵測",
        ],
        "note": "CSI = H(f)：頻域通道矩陣，每對 TX-RX 天線各一組",
    },
    # Slide 5: MAC 層職責
    {
        "type": "bullets",
        "title": "MAC 層職責",
        "bullets": [
            "媒介存取控制：CSMA/CA（避免碰撞）",
            "確認機制：ACK frame",
            "Frame 組裝：802.11 Header + Payload",
            "安全：WPA2/3 加解密（CCMP/GCMP）",
            "省電機制：PS-Poll、TWT（Wi-Fi 6）",
            "",
            "▶ CSI 在 MAC 層的狀態",
            "  · PHY 解調完成後 CSI 儲存於 NIC 韌體/暫存器",
            "  · MAC 層不負責傳遞 CSI——它只處理 frame 內容",
            "  · CSI 屬於 PHY 層 side-channel，不在 802.11 協定定義範圍內",
        ],
    },
    # Slide 6: OSI 7 層 vs Windows 元件
    {
        "type": "table",
        "title": "OSI 7 層 vs Windows 實際元件對應",
        "headers": ["OSI 層", "標準名稱", "Windows 元件", "備註"],
        "rows": [
            ["L7", "Application", "應用程式 / WinSock API", "send() / recv()"],
            ["L4", "Transport", "tcpip.sys (TCP/UDP)", "連接埠管理"],
            ["L3", "Network", "tcpip.sys (IP)", "路由、IP 封包"],
            ["L2b", "LLC", "ndis.sys (NDIS)", "協定多工/解多工"],
            ["L2a", "MAC", "nwifi.sys + OEM.sys", "802.11 管理"],
            ["L1", "PHY", "NIC 韌體 + 硬體", "RF 信號處理，CSI 在此層"],
        ],
        "note": "CSI 存在於 L1，NDIS 介面從 L2a 開始，兩者之間沒有標準橋接",
    },
    # Slide 7: Windows 驅動堆疊總覽
    {
        "type": "stack_diagram",
        "title": "Windows 網路驅動架構總覽",
        "layers": [
            ("應用程式", "User Mode", "#2d5a27"),
            ("WinSock (ws2_32.dll)", "User Mode", "#2d5a27"),
            ("AFD.sys（Ancillary Function Driver）", "Kernel Mode", "#1a3a6b"),
            ("tcpip.sys（TCP/IP Stack）", "Kernel Mode", "#1a3a6b"),
            ("ndis.sys（NDIS）", "Kernel Mode", "#4a2080"),
            ("nwifi.sys（Native WiFi）", "Kernel Mode", "#4a2080"),
            ("OEM.sys（Miniport Driver）", "Kernel Mode", "#7a1a1a"),
            ("NIC 硬體 + 韌體", "Hardware", "#3a3a00"),
        ],
        "note": "每層只能透過明確定義的介面與相鄰層溝通",
    },
    # Slide 8: OEM.sys 是什麼
    {
        "type": "two_col",
        "title": "OEM.sys：網卡廠商 Miniport Driver",
        "left_title": "職責",
        "left_bullets": [
            "直接與 NIC 硬體暫存器溝通（MMIO/PCI）",
            "控制 RF 參數（頻道、頻寬、TX 功率）",
            "管理 DMA 傳輸（TX/RX 封包緩衝）",
            "實作 NDIS Miniport 介面",
            "上報硬體事件給 nwifi.sys",
        ],
        "right_title": "範例（各廠商）",
        "right_bullets": [
            "Intel：iwlwifi（Linux）/ netwtw*.sys（Windows）",
            "Qualcomm/Atheros：ath*.sys",
            "Broadcom：bcmwl*.sys",
            "MediaTek：mt*.sys",
            "",
            "⚠ OEM.sys 是唯一能讀取 CSI 的軟體元件",
            "  但它選擇不向上層暴露此資料",
        ],
        "note": "OEM.sys 是 Windows 驅動堆疊中距離硬體最近的軟體層",
    },
    # Slide 9: OEM.sys 資料路徑
    {
        "type": "flow",
        "title": "OEM.sys 的資料路徑",
        "flow_items": [
            ("NIC 韌體\n計算 CSI", "儲存於暫存器"),
            ("OEM.sys\nDMA 讀取", "RX 封包緩衝"),
            ("NDIS_RECEIVE\n_NET_BUFFER", "標準封包上報"),
            ("nwifi.sys\n接收", "802.11 frame 處理"),
            ("tcpip.sys\n→ 應用層", "IP 封包"),
        ],
        "note": "CSI 存於 NIC 暫存器，OEM.sys 只將 802.11 frame payload 上報——CSI 從未進入 NDIS 封包結構",
    },
    # Slide 10: nwifi.sys 是什麼
    {
        "type": "bullets",
        "title": "nwifi.sys：Windows Native WiFi 驅動",
        "bullets": [
            "位置：位於 OEM.sys 上方，ndis.sys 下方",
            "",
            "主要職責：",
            "  · 802.11 認證流程（Open / WPA2 / WPA3）",
            "  · 漫遊管理（BSS 選擇、重新關聯）",
            "  · 省電模式（Legacy PS / U-APSD）",
            "  · 掃描管理（Passive / Active Scan）",
            "  · 802.11 管理幀處理（Probe / Auth / Assoc）",
            "",
            "暴露給上層的介面：",
            "  · Native WiFi API（wlanapi.dll）→ 應用程式",
            "  · NDIS OID 介面 → tcpip.sys",
            "",
            "⚠ nwifi.sys 不處理、不儲存、不轉發任何 PHY 層原始資料",
        ],
    },
    # Slide 11: nwifi.sys 抽象化行為
    {
        "type": "two_col",
        "title": "nwifi.sys 的抽象化：CSI 在此消失",
        "left_title": "nwifi.sys 向上暴露的資訊",
        "left_bullets": [
            "SSID / BSSID",
            "RSSI（接收信號強度）",
            "連線狀態（已連線/斷線）",
            "安全性設定（加密方式）",
            "頻道 / 頻寬",
            "連線速率（Link Speed）",
        ],
        "right_title": "nwifi.sys 不暴露的資訊",
        "right_bullets": [
            "❌ CSI（子載波增益矩陣）",
            "❌ 原始 IQ 樣本",
            "❌ 每根天線的原始 RSSI",
            "❌ 噪聲圖（Noise Floor per subcarrier）",
            "❌ 封包時間戳（硬體層級）",
            "",
            "設計原因：nwifi.sys 的目標是",
            "連線管理，不是 RF 量測",
        ],
        "note": "從 nwifi.sys 向上，PHY 層資訊已被永久丟棄",
    },
    # Slide 12: NDIS 設計哲學
    {
        "type": "bullets",
        "title": "NDIS 介面設計哲學",
        "bullets": [
            "NDIS = Network Driver Interface Specification",
            "設計目標：讓任何網路協定可搭配任何網卡——跨廠商抽象化",
            "",
            "OID（Object Identifier）機制：",
            "  · OID_802_11_SSID           → 查詢/設定 SSID",
            "  · OID_802_11_RSSI           → 查詢接收信號強度",
            "  · OID_802_11_BSSID_LIST     → 掃描結果",
            "  · OID_802_11_STATISTICS     → 封包統計",
            "",
            "  ❌ 無任何 OID 定義 CSI 查詢",
            "  ❌ 無任何 OID 定義子載波層級資訊",
            "",
            "NDIS 的「通用性」正是 CSI 無法取得的根本原因：",
            "標準化需要取最大公約數，CSI 是廠商私有的 PHY 細節",
        ],
    },
    # Slide 13: WDI
    {
        "type": "two_col",
        "title": "WDI：Windows 10+ 新驅動模型",
        "left_title": "WDI 改變了什麼",
        "left_bullets": [
            "WLAN Device Driver Interface（WDI）",
            "Win10 引入，取代部分 NDIS WiFi 介面",
            "將驅動分為：",
            "  · IHV 元件（UMDF，User Mode）",
            "  · WDI 框架（KMDF，Kernel Mode）",
            "降低藍屏風險（User Mode crash 不致命）",
            "簡化 IHV 驅動開發",
        ],
        "right_title": "WDI 對 CSI 的影響",
        "right_bullets": [
            "❌ WDI Task/Property 定義中",
            "   同樣沒有 CSI 相關介面",
            "",
            "WDITASK 清單（部分）：",
            "  · TaskScan",
            "  · TaskConnect",
            "  · TaskDisconnect",
            "  · TaskSendRequest",
            "",
            "→ 架構更新，但 CSI 缺席問題未解決",
        ],
        "note": "WDI 是架構改進，不是 PHY 資料暴露的嘗試",
    },
    # Slide 14: CSI 消失點分析
    {
        "type": "stack_diagram_annotated",
        "title": "CSI 消失點分析",
        "layers": [
            ("應用程式 / WinSock", "無法感知 CSI 存在", "normal"),
            ("tcpip.sys", "只處理 IP 封包", "normal"),
            ("ndis.sys（NDIS）", "無 CSI OID 定義", "warning"),
            ("nwifi.sys", "丟棄所有 PHY metadata", "danger"),
            ("OEM.sys", "能讀取 CSI，但不上報", "danger"),
            ("NIC 韌體 + 硬體", "CSI 計算並儲存於此", "source"),
        ],
        "note": "CSI 消失於 OEM.sys ↔ nwifi.sys 介面：OEM.sys 有能力讀取但選擇不轉發，nwifi.sys 也不要求",
    },
    # Slide 15: 為什麼 Linux 可以
    {
        "type": "two_col",
        "title": "為什麼 Linux 可以取得 CSI？",
        "left_title": "Linux 驅動架構",
        "left_bullets": [
            "cfg80211：802.11 設定框架",
            "mac80211：軟體 MAC 實作",
            "nl80211：Netlink 介面（用戶空間↔核心）",
            "",
            "關鍵差異：開源驅動",
            "  · iwlwifi（Intel）可修改",
            "  · ath9k（Atheros）可修改",
            "  · brcmfmac（Broadcom）可 patch 韌體",
            "",
            "debugfs 介面：",
            "  /sys/kernel/debug/ieee80211/",
            "  部分驅動在此暴露 CSI 原始資料",
        ],
        "right_title": "Windows 的根本差異",
        "right_bullets": [
            "✗ OEM.sys 為閉源二進位",
            "✗ NDIS 介面固定，無擴充 CSI 的機制",
            "✗ 無等同 debugfs 的 PHY 資料介面",
            "✗ 韌體修改需要 WHQL 簽章",
            "",
            "→ Linux 的可行性來自",
            "  開源 + 可修改驅動 + 彈性的核心介面",
            "  不是 Windows 的「bug」，是設計哲學差異",
        ],
    },
    # Slide 16: linux-80211n-csitool
    {
        "type": "bullets",
        "title": "Linux CSI Tool：linux-80211n-csitool（Intel 5300）",
        "bullets": [
            "作者：Daniel Halperin（University of Washington，2011）",
            "硬體：Intel WiFi Link 5300（IWL5300），3 天線",
            "",
            "運作原理：",
            "  1. 修改 Intel 官方韌體（firmware patch）",
            "  2. 自訂 iwlwifi 核心驅動",
            "  3. 每個 RX 封包附帶 CSI 矩陣（30 個子載波群組）",
            "  4. 每個子載波：複數值（8-bit real + 8-bit imag）",
            "",
            "平台：Ubuntu 10.04 LTS + kernel 2.6.36",
            "  ⚠ 歷史性工具，現代 Linux 核心需額外移植工作",
            "",
            "資料格式：每個 RX 封包 → CSI 矩陣 [Ntx × Nrx × 30]",
            "GitHub: dhalperi/linux-80211n-csitool",
        ],
    },
    # Slide 17: 現代 CSI 工具生態
    {
        "type": "table",
        "title": "現代 CSI 工具生態（全部 Linux-only）",
        "headers": ["工具", "硬體", "標準", "頻寬", "特色"],
        "rows": [
            ["PicoScenes", "Intel AX200/AX210", "802.11a/g/n/ac/ax", "20–160 MHz", "最完整，商業支援"],
            ["FeitCSI", "Intel AX200/AX210", "802.11a/g/n/ac/ax", "20–160 MHz", "開源，CSI 注入"],
            ["IAX", "Intel AX200/201/210/211", "802.11ax", "20–160 MHz", "學術研究用"],
            ["Nexmon CSI", "Broadcom/Cypress", "802.11a/g/n/ac", "20–80 MHz", "Raspberry Pi 支援"],
            ["Atheros-CSI-Tool", "AR9380（ath9k）", "802.11n", "20/40 MHz", "OpenWRT 支援"],
            ["linux-80211n-csitool", "Intel IWL5300", "802.11n", "20/40 MHz", "最早的工具，歷史性"],
        ],
        "note": "所有工具均僅支援 Linux——這直接說明了為什麼 Windows 無法取得 CSI",
    },
    # Slide 18: Windows 替代方案 1
    {
        "type": "bullets",
        "title": "Windows 替代方案 1：Raw Packet Capture",
        "bullets": [
            "工具：WinPcap / Npcap + Monitor Mode（若網卡支援）",
            "",
            "可以取得：",
            "  ✅ 802.11 frame header（MAC Header）",
            "  ✅ Radiotap Header（部分 PHY 資訊）",
            "  ✅ RSSI（訊號強度，dBm）",
            "  ✅ 資料速率（MCS index）",
            "  ✅ 頻道資訊",
            "",
            "無法取得：",
            "  ❌ CSI（子載波增益矩陣）",
            "  ❌ IQ 樣本",
            "  ❌ 多天線個別 RSSI",
            "",
            "適用情境：封包分析、協定研究、基本 RF 環境監控",
            "限制：Monitor Mode 在 Windows 支援有限，需特定網卡與驅動版本",
        ],
    },
    # Slide 19: Windows 替代方案 2
    {
        "type": "bullets",
        "title": "Windows 替代方案 2：OEM 私有 IOCTL",
        "bullets": [
            "原理：部分廠商在 OEM.sys 中保留私有控制介面",
            "透過 DeviceIoControl() 呼叫私有 IOCTL code",
            "",
            "需要的工作：",
            "  1. 逆向工程 OEM.sys（IDA Pro / Ghidra）",
            "  2. 找到處理 CSI 的 IOCTL dispatch 函式",
            "  3. 確認輸入/輸出緩衝區格式",
            "  4. 撰寫呼叫程式（需以管理員權限執行）",
            "",
            "現實情況：",
            "  · Intel：部分版本 netwtw*.sys 有私有介面，但未公開文件",
            "  · 每次驅動更新後 IOCTL code 可能改變",
            "  · 可能觸發驅動簽章驗證失敗（Win10 Secure Boot）",
            "",
            "⚠ 高風險、高維護成本，不建議用於生產環境",
        ],
    },
    # Slide 20: Windows 替代方案 3
    {
        "type": "bullets",
        "title": "Windows 替代方案 3：雙系統 / 虛擬機",
        "bullets": [
            "最實用的工程折衷方案",
            "",
            "架構 A：雙系統（推薦）",
            "  · 同一台機器安裝 Linux + Windows",
            "  · Linux 環境：使用 PicoScenes / FeitCSI 取 CSI",
            "  · CSI 資料儲存為 .mat / .csv / .bin",
            "  · 重開機進 Windows 進行後處理（MATLAB / Python）",
            "",
            "架構 B：USB Live Linux",
            "  · 免安裝，開機從 USB 進 Linux",
            "  · FeitCSI 提供 Live USB 映像",
            "  · 完成採集後重開機回 Windows",
            "",
            "架構 C：VM + USB 直通（不推薦）",
            "  · WiFi 網卡 USB 直通至 Linux VM",
            "  · 延遲與時序精度受 Hypervisor 影響",
            "  · PCIe 網卡通常無法直通",
            "",
            "推薦：雙系統 + 共享資料分割區",
        ],
    },
    # Slide 21: 替代方案比較表
    {
        "type": "table",
        "title": "替代方案比較",
        "headers": ["方案", "技術難度", "硬體需求", "CSI 完整度", "維護成本"],
        "rows": [
            ["Raw Packet Capture", "★☆☆☆☆", "一般 WiFi 網卡", "❌ 無 CSI", "低"],
            ["OEM 私有 IOCTL", "★★★★☆", "特定廠商網卡", "⚠ 部分，不穩定", "極高"],
            ["雙系統（推薦）", "★★☆☆☆", "支援 CSI 的網卡", "✅ 完整", "低"],
            ["VM + USB 直通", "★★★☆☆", "USB WiFi 網卡", "⚠ 時序精度差", "中"],
            ["Linux Live USB", "★☆☆☆☆", "支援 CSI 的網卡", "✅ 完整", "低"],
        ],
        "note": "對 RF 工程師的建議：雙系統 + PicoScenes/FeitCSI 是最佳投資報酬比",
    },
    # Slide 22: 硬體選型
    {
        "type": "table",
        "title": "支援 CSI 的硬體選型",
        "headers": ["網卡型號", "晶片", "標準", "CSI 工具", "平台"],
        "rows": [
            ["Intel WiFi Link 5300", "IWL5300", "802.11n", "linux-80211n-csitool", "Linux（舊）"],
            ["Intel Wi-Fi 6 AX200", "AX200", "802.11ax", "PicoScenes / FeitCSI", "Linux"],
            ["Intel Wi-Fi 6E AX210", "AX210", "802.11ax（6GHz）", "PicoScenes / FeitCSI", "Linux"],
            ["Atheros AR9380", "AR9380", "802.11n", "Atheros-CSI-Tool", "Linux / OpenWRT"],
            ["Broadcom BCM4339", "BCM4339", "802.11ac", "Nexmon CSI", "Android / RPi"],
            ["Raspberry Pi 4 (RPi)", "BCM43455", "802.11ac", "Nexmon CSI", "Linux（RPi OS）"],
        ],
        "note": "採購建議：Intel AX200 M.2 網卡 + Linux 雙系統 = 最易取得的 CSI 研究平台",
    },
    # Slide 23: 架構限制總結
    {
        "type": "stack_diagram_annotated",
        "title": "完整架構總覽：CSI 的旅程",
        "layers": [
            ("應用程式", "只能用 WinSock，看不到 PHY", "normal"),
            ("NDIS / tcpip.sys", "無 CSI OID，協定通用化", "warning"),
            ("nwifi.sys", "802.11 管理，丟棄 PHY metadata", "danger"),
            ("OEM.sys（唯一機會）", "能讀 CSI，但介面不暴露", "danger"),
            ("NIC 韌體", "CSI 在此計算完成", "source"),
            ("PHY 硬體 + RF 電路", "OFDM 解調，子載波分析", "source"),
        ],
        "note": "Windows 架構的每一層都有合理的設計動機，但合力造成了 CSI 無法取得的結果",
    },
    # Slide 24: 給 RF 工程師的建議
    {
        "type": "bullets",
        "title": "給 RF 工程師的建議",
        "bullets": [
            "1. 選硬體前先確認 CSI 工具支援",
            "   → Intel AX200/AX210 是目前最佳選擇",
            "",
            "2. 優先使用 Linux 環境採集 CSI",
            "   → PicoScenes（商業支援）或 FeitCSI（開源）",
            "   → 雙系統方案最穩定",
            "",
            "3. Windows 定位為後處理平台",
            "   → MATLAB / Python 分析 CSI 資料",
            "   → 不要嘗試在 Windows 直接取 CSI",
            "",
            "4. 避免的坑",
            "   ❌ 不要依賴 OEM 私有 IOCTL（維護成本極高）",
            "   ❌ 不要用 VM 取 CSI（時序不準確）",
            "   ❌ 不要購買沒有明確工具支援的網卡",
            "",
            "5. 資源",
            "   · PicoScenes: ps.zpj.io",
            "   · FeitCSI: feitcsi.kuskosoft.com",
            "   · CSIKit（多格式解析）: github.com/Gi-z/CSIKit",
        ],
    },
    # Slide 25: Q&A / 參考資料
    {
        "type": "references",
        "title": "參考資料",
        "refs": [
            "[1] D. Halperin et al., "Tool Release: Gathering 802.11n Traces with CSI," ACM SIGCOMM, 2011. → linux-80211n-csitool",
            "[2] M. Schulz et al., "Nexmon: The C-based Firmware Patching Framework," — seemoo-lab/nexmon_csi",
            "[3] Z. Jiang, "PicoScenes: Enabling Modern Wi-Fi ISAC Research," — ps.zpj.io",
            "[4] KuskoSoft, "FeitCSI: 802.11 CSI Tool," — feitcsi.kuskosoft.com",
            "[5] X. Yaxiong et al., "Atheros CSI Tool for 802.11n NICs," — xieyaxiongfly/Atheros-CSI-Tool",
            "[6] Microsoft Docs, "WDI Miniport Driver Design Guide," — learn.microsoft.com",
            "[7] Microsoft Docs, "NDIS Network Interface Architecture," — learn.microsoft.com",
            "[8] Gi-z, "CSIKit: Python CSI Processing Tools," — github.com/Gi-z/CSIKit",
        ],
        "qanda": "Q&A",
    },
]
```

- [ ] **Step 2: 驗證語法**

```bash
python -c "import sys; sys.path.insert(0,'src'); import content; print(f'Slides defined: {len(content.SLIDES)}')"
```

Expected: `Slides defined: 25`

---

## Task 4：建立 builders.py

**Files:**
- Create: `src/builders.py`

- [ ] **Step 1: 建立投影片版型函式**

```python
# src/builders.py
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import theme

def set_slide_background(slide, color: RGBColor):
    """設定投影片背景色"""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_textbox(slide, text, left, top, width, height,
                font_name=theme.FONT_BODY, font_size=theme.BODY_SIZE,
                color=theme.TEXT_COLOR, bold=False, align=PP_ALIGN.LEFT,
                word_wrap=True):
    """新增文字方塊"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = font_size
    run.font.color.rgb = color
    run.font.bold = bold
    return txBox

def add_title(slide, title_text):
    """新增標準標題列"""
    left = Inches(0.4)
    top = Inches(0.2)
    width = Inches(9.2)
    height = Inches(0.8)
    add_textbox(slide, title_text, left, top, width, height,
                font_name=theme.FONT_TITLE, font_size=theme.TITLE_SIZE,
                color=theme.PRIMARY_COLOR, bold=True)
    # 標題下方分隔線
    line = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.LINE
        left, Inches(1.0), width, Emu(0)
    )
    line.line.color.rgb = theme.PRIMARY_COLOR
    line.line.width = Pt(1.5)

def add_note(slide, note_text):
    """底部備註列"""
    left = Inches(0.4)
    top = Inches(6.8)
    width = Inches(9.2)
    height = Inches(0.5)
    add_textbox(slide, f"▶ {note_text}", left, top, width, height,
                font_size=theme.SMALL_SIZE, color=theme.SUBTEXT_COLOR)

def build_cover(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    # 主標題
    add_textbox(slide, data["title"],
                Inches(0.5), Inches(1.5), Inches(9.0), Inches(1.5),
                font_name=theme.FONT_TITLE, font_size=Pt(44),
                color=theme.PRIMARY_COLOR, bold=True, align=PP_ALIGN.CENTER)
    # 副標
    add_textbox(slide, data["subtitle"],
                Inches(0.5), Inches(3.0), Inches(9.0), Inches(1.2),
                font_size=theme.SUBTITLE_SIZE,
                color=theme.SUBTEXT_COLOR, align=PP_ALIGN.CENTER)
    # 日期
    add_textbox(slide, data["date"],
                Inches(0.5), Inches(6.5), Inches(9.0), Inches(0.5),
                font_size=theme.SMALL_SIZE,
                color=theme.SUBTEXT_COLOR, align=PP_ALIGN.CENTER)

def build_bullets(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    add_title(slide, data["title"])
    left = Inches(0.5)
    top = Inches(1.2)
    width = Inches(9.0)
    height = Inches(5.5)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for bullet in data["bullets"]:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        run.text = bullet
        if bullet.startswith("  "):
            run.font.size = Pt(16)
            run.font.color.rgb = theme.SUBTEXT_COLOR
            p.level = 1
        elif bullet == "":
            run.font.size = Pt(8)
        else:
            run.font.size = theme.BODY_SIZE
            run.font.color.rgb = theme.TEXT_COLOR
            if bullet.startswith("▶") or bullet.endswith("：") or bullet[0].isdigit():
                run.font.bold = True
                run.font.color.rgb = theme.PRIMARY_COLOR
    if "note" in data:
        add_note(slide, data["note"])

def build_two_col(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    add_title(slide, data["title"])

    def add_col(title, bullets, left):
        # 欄標題
        add_textbox(slide, title, left, Inches(1.2), Inches(4.3), Inches(0.5),
                    font_size=Pt(18), color=theme.ACCENT_COLOR, bold=True)
        # 欄內容
        txBox = slide.shapes.add_textbox(left, Inches(1.8), Inches(4.3), Inches(4.8))
        tf = txBox.text_frame
        tf.word_wrap = True
        first = True
        for b in bullets:
            if first:
                p = tf.paragraphs[0]
                first = False
            else:
                p = tf.add_paragraph()
            run = p.add_run()
            run.text = b
            if b.startswith("  "):
                run.font.size = Pt(14)
                run.font.color.rgb = theme.SUBTEXT_COLOR
            elif b == "":
                run.font.size = Pt(6)
            else:
                run.font.size = Pt(16)
                run.font.color.rgb = theme.TEXT_COLOR
                if b.startswith("✅") or b.startswith("❌") or b.startswith("⚠"):
                    run.font.bold = True

    add_col(data["left_title"], data["left_bullets"], Inches(0.3))
    # 分隔線
    div = slide.shapes.add_shape(1, Inches(4.8), Inches(1.1), Emu(0), Inches(5.5))
    div.line.color.rgb = theme.PRIMARY_COLOR
    div.line.width = Pt(0.75)
    add_col(data["right_title"], data["right_bullets"], Inches(5.0))
    if "note" in data:
        add_note(slide, data["note"])

def build_table(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    add_title(slide, data["title"])
    rows = len(data["rows"]) + 1
    cols = len(data["headers"])
    left = Inches(0.3)
    top = Inches(1.2)
    width = Inches(9.4)
    height = Inches(0.4 * rows + 0.1)
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    col_width = int(width / cols)
    for i in range(cols):
        table.columns[i].width = col_width
    # 表頭
    for ci, hdr in enumerate(data["headers"]):
        cell = table.cell(0, ci)
        cell.text = hdr
        cell.fill.solid()
        cell.fill.fore_color.rgb = theme.TABLE_HDR_BG
        p = cell.text_frame.paragraphs[0]
        run = p.runs[0] if p.runs else p.add_run()
        run.font.bold = True
        run.font.size = Pt(15)
        run.font.color.rgb = theme.TEXT_COLOR
        run.font.name = theme.FONT_BODY
        p.alignment = PP_ALIGN.CENTER
    # 資料列
    for ri, row in enumerate(data["rows"]):
        bg = theme.BG_COLOR if ri % 2 == 0 else theme.TABLE_ROW_ALT
        for ci, val in enumerate(row):
            cell = table.cell(ri + 1, ci)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            run = p.runs[0] if p.runs else p.add_run()
            run.font.size = Pt(13)
            run.font.color.rgb = theme.TEXT_COLOR
            run.font.name = theme.FONT_BODY
    if "note" in data:
        add_note(slide, data["note"])

def build_flow(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    add_title(slide, data["title"])
    items = data["flow_items"]
    n = len(items)
    box_w = Inches(1.5)
    box_h = Inches(0.9)
    gap = Inches(0.2)
    total_w = n * box_w + (n - 1) * gap
    start_x = (Inches(10) - total_w) / 2
    y = Inches(2.5)
    for i, (label, desc) in enumerate(items):
        x = start_x + i * (box_w + gap)
        shape = slide.shapes.add_shape(
            1, x, y, box_w, box_h  # rectangle
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(0x0d, 0x47, 0xa1)
        shape.line.color.rgb = theme.PRIMARY_COLOR
        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = label
        run.font.size = Pt(13)
        run.font.color.rgb = theme.TEXT_COLOR
        run.font.bold = True
        # 描述文字
        add_textbox(slide, desc, x, y + box_h + Inches(0.1),
                    box_w, Inches(0.4),
                    font_size=Pt(11), color=theme.SUBTEXT_COLOR,
                    align=PP_ALIGN.CENTER)
        # 箭頭
        if i < n - 1:
            arr_x = x + box_w
            arr = slide.shapes.add_shape(1, arr_x, y + box_h/2 - Pt(3), gap, Pt(6))
            arr.fill.solid()
            arr.fill.fore_color.rgb = theme.PRIMARY_COLOR
            arr.line.color.rgb = theme.PRIMARY_COLOR
    if "note" in data:
        add_note(slide, data["note"])

def build_stack_diagram(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    add_title(slide, data["title"])
    layers = data["layers"]
    n = len(layers)
    box_h = Inches(0.55)
    gap = Inches(0.05)
    total_h = n * (box_h + gap)
    start_y = (Inches(6.5) - total_h) / 2 + Inches(0.8)
    color_map = {
        "#2d5a27": RGBColor(0x2d, 0x5a, 0x27),
        "#1a3a6b": RGBColor(0x1a, 0x3a, 0x6b),
        "#4a2080": RGBColor(0x4a, 0x20, 0x80),
        "#7a1a1a": RGBColor(0x7a, 0x1a, 0x1a),
        "#3a3a00": RGBColor(0x3a, 0x3a, 0x00),
    }
    for i, layer_info in enumerate(layers):
        if len(layer_info) == 3:
            label, sublabel, color_hex = layer_info
        else:
            label, sublabel = layer_info
            color_hex = "#1a3a6b"
        y = start_y + i * (box_h + gap)
        fill_color = color_map.get(color_hex, RGBColor(0x1a, 0x3a, 0x6b))
        shape = slide.shapes.add_shape(1, Inches(1.0), y, Inches(8.0), box_h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
        shape.line.color.rgb = theme.PRIMARY_COLOR
        tf = shape.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = label
        run.font.size = Pt(16)
        run.font.color.rgb = theme.TEXT_COLOR
        run.font.bold = True
        if sublabel:
            add_textbox(slide, sublabel,
                        Inches(9.2), y, Inches(0.8), box_h,
                        font_size=Pt(10), color=theme.SUBTEXT_COLOR)
    if "note" in data:
        add_note(slide, data["note"])

def build_stack_diagram_annotated(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    add_title(slide, data["title"])
    layers = data["layers"]
    n = len(layers)
    box_h = Inches(0.58)
    gap = Inches(0.06)
    total_h = n * (box_h + gap)
    start_y = Inches(1.1)
    status_colors = {
        "normal":  RGBColor(0x1a, 0x3a, 0x6b),
        "warning": RGBColor(0x6b, 0x4a, 0x00),
        "danger":  RGBColor(0x6b, 0x1a, 0x1a),
        "source":  RGBColor(0x1a, 0x5a, 0x27),
    }
    for i, (label, annotation, status) in enumerate(layers):
        y = start_y + i * (box_h + gap)
        fill_color = status_colors.get(status, status_colors["normal"])
        shape = slide.shapes.add_shape(1, Inches(0.3), y, Inches(5.5), box_h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
        shape.line.color.rgb = theme.PRIMARY_COLOR
        tf = shape.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = label
        run.font.size = Pt(15)
        run.font.color.rgb = theme.TEXT_COLOR
        run.font.bold = True
        # 標注
        ann_color = theme.ACCENT_COLOR if status in ("danger", "warning") else theme.SUBTEXT_COLOR
        if status == "source":
            ann_color = RGBColor(0x66, 0xff, 0x66)
        add_textbox(slide, annotation,
                    Inches(6.0), y + Inches(0.1), Inches(3.8), box_h,
                    font_size=Pt(14), color=ann_color)
    if "note" in data:
        add_note(slide, data["note"])

def build_references(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    add_title(slide, data.get("qanda", "Q&A") + " / 參考資料")
    txBox = slide.shapes.add_textbox(Inches(0.4), Inches(1.2), Inches(9.2), Inches(5.0))
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for ref in data["refs"]:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        run.text = ref
        run.font.size = Pt(13)
        run.font.color.rgb = theme.SUBTEXT_COLOR
        run.font.name = theme.FONT_BODY

BUILDERS = {
    "cover":                    build_cover,
    "bullets":                  build_bullets,
    "two_col":                  build_two_col,
    "table":                    build_table,
    "flow":                     build_flow,
    "stack_diagram":            build_stack_diagram,
    "stack_diagram_annotated":  build_stack_diagram_annotated,
    "references":               build_references,
}
```

- [ ] **Step 2: 驗證 import**

```bash
python -c "import sys; sys.path.insert(0,'src'); import builders; print('builders OK, types:', list(builders.BUILDERS.keys()))"
```

Expected: `builders OK, types: ['cover', 'bullets', 'two_col', 'table', 'flow', 'stack_diagram', 'stack_diagram_annotated', 'references']`

---

## Task 5：建立 create_ppt.py

**Files:**
- Create: `src/create_ppt.py`

- [ ] **Step 1: 建立主程式**

```python
# src/create_ppt.py
import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

from pptx import Presentation
from pptx.util import Emu
import theme
import content
import builders

OUTPUT_PATH = os.path.join(
    os.path.dirname(__file__), "..", "output", "network-card-csi.pptx"
)

def create_presentation():
    prs = Presentation()
    prs.slide_width = theme.SLIDE_WIDTH
    prs.slide_height = theme.SLIDE_HEIGHT

    blank_layout = prs.slide_layouts[6]  # 完全空白版型

    for slide_data in content.SLIDES:
        slide = prs.slides.add_slide(blank_layout)
        slide_type = slide_data["type"]
        builder_fn = builders.BUILDERS.get(slide_type)
        if builder_fn is None:
            print(f"WARNING: unknown slide type '{slide_type}', skipping")
            continue
        builder_fn(slide, slide_data)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    prs.save(OUTPUT_PATH)
    print(f"Saved: {OUTPUT_PATH}")
    print(f"Total slides: {len(prs.slides)}")

if __name__ == "__main__":
    create_presentation()
```

- [ ] **Step 2: 執行並驗證輸出**

```bash
cd /c/Users/Jacky/Desktop/network-ppt-project && python src/create_ppt.py
```

Expected:
```
Saved: .../output/network-card-csi.pptx
Total slides: 25
```

- [ ] **Step 3: 確認檔案存在**

```bash
ls -la /c/Users/Jacky/Desktop/network-ppt-project/output/
```

Expected: `network-card-csi.pptx` 存在且大小 > 50KB

---

## Task 6：撰寫測試並驗證

**Files:**
- Create: `tests/test_ppt.py`

- [ ] **Step 1: 撰寫測試**

```python
# tests/test_ppt.py
import sys
import os
import pytest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from pptx import Presentation
from pptx.dml.color import RGBColor
import theme

PPT_PATH = os.path.join(os.path.dirname(__file__), '..', 'output', 'network-card-csi.pptx')

EXPECTED_TITLES = [
    "網卡與 OS 層的關係",  # 封面
    "大綱",
    "802.11 RF 信號路徑",
    "PHY 層：什麼是 CSI？",
    "MAC 層職責",
    "OSI 7 層 vs Windows 實際元件對應",
    "Windows 網路驅動架構總覽",
    "OEM.sys：網卡廠商 Miniport Driver",
    "OEM.sys 的資料路徑",
    "nwifi.sys：Windows Native WiFi 驅動",
    "nwifi.sys 的抽象化：CSI 在此消失",
    "NDIS 介面設計哲學",
    "WDI：Windows 10+ 新驅動模型",
    "CSI 消失點分析",
    "為什麼 Linux 可以取得 CSI？",
    "Linux CSI Tool：linux-80211n-csitool（Intel 5300）",
    "現代 CSI 工具生態（全部 Linux-only）",
    "Windows 替代方案 1：Raw Packet Capture",
    "Windows 替代方案 2：OEM 私有 IOCTL",
    "Windows 替代方案 3：雙系統 / 虛擬機",
    "替代方案比較",
    "支援 CSI 的硬體選型",
    "完整架構總覽：CSI 的旅程",
    "給 RF 工程師的建議",
    "參考資料",
]

@pytest.fixture(scope="module")
def prs():
    assert os.path.exists(PPT_PATH), f"PPT not found: {PPT_PATH}"
    return Presentation(PPT_PATH)

def test_slide_count(prs):
    assert len(prs.slides) == 25

def test_slide_dimensions(prs):
    assert prs.slide_width == theme.SLIDE_WIDTH
    assert prs.slide_height == theme.SLIDE_HEIGHT

def test_slide_17_has_table(prs):
    slide = prs.slides[16]  # 0-indexed
    tables = [s for s in slide.shapes if s.has_table]
    assert len(tables) >= 1, "Slide 17 should have a table"

def test_slide_21_has_table(prs):
    slide = prs.slides[20]
    tables = [s for s in slide.shapes if s.has_table]
    assert len(tables) >= 1, "Slide 21 should have a table"

def test_slide_22_has_table(prs):
    slide = prs.slides[21]
    tables = [s for s in slide.shapes if s.has_table]
    assert len(tables) >= 1, "Slide 22 should have a table"

def test_all_slides_have_shapes(prs):
    for i, slide in enumerate(prs.slides):
        assert len(slide.shapes) > 0, f"Slide {i+1} has no shapes"
```

- [ ] **Step 2: 安裝 pytest 並執行**

```bash
pip install pytest && cd /c/Users/Jacky/Desktop/network-ppt-project && python -m pytest tests/test_ppt.py -v
```

Expected: 所有測試 PASSED

---

## Task 7：最終確認

- [ ] **Step 1: 用 PowerPoint / LibreOffice 開啟確認**

```bash
explorer.exe "C:\\Users\\Jacky\\Desktop\\network-ppt-project\\output\\network-card-csi.pptx"
```

- [ ] **Step 2: 目視確認以下重點投影片**
- Slide 6：有兩欄對比表格
- Slide 7：有堆疊方塊圖
- Slide 14：有標注顏色的堆疊圖
- Slide 17：有工具比較表
- Slide 21：有替代方案比較表

- [ ] **Step 3: 完成**

輸出位置：`C:/Users/Jacky/Desktop/network-ppt-project/output/network-card-csi.pptx`
