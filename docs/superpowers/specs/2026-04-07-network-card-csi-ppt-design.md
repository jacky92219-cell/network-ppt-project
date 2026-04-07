# 設計文件：網卡與 OS 層關係 PPT

**日期：** 2026-04-07
**主題：** 網卡與 OS 層的關係 — 為什麼 Windows 無法取得網卡 CSI
**受眾：** RF 工程師
**格式：** 25 張投影片，標準版（45–60 分鐘）
**製作工具：** python-pptx

---

## 目標

1. 教學：讓 RF 工程師理解 Windows 網路驅動堆疊如何與 RF 硬體溝通
2. 技術分享：內部 team 同步架構知識
3. 開發參考：幫助 RF 工程師理解 CSI 資料如何從硬體傳至 OS，以及為什麼在 Windows 上取不到

---

## 結構方案

採用**由下而上（Bottom-Up）**方式，從 RF 工程師熟悉的 PHY/MAC 層出發，逐層向上說明，最後揭示 CSI 被截斷的位置與替代方案。

---

## 投影片結構

### 第一段：基礎架構（Slides 1–6）

| # | 標題 | 內容重點 |
|---|------|---------|
| 1 | 封面 | 標題、副標、作者 |
| 2 | 大綱 | 全程路線圖 |
| 3 | 802.11 RF 信號路徑 | 天線 → ADC → OFDM 解調 → 802.11 frame |
| 4 | PHY 層：什麼是 CSI？ | Channel State Information 定義、OFDM 子載波振幅/相位、與 RSSI 的差異 |
| 5 | MAC 層職責 | CSMA/CA、ACK、frame 組裝，CSI 在 MAC 層的狀態 |
| 6 | OSI 7 層 vs Windows 實際元件對應 | L1=PHY、L2=MAC/OEM.sys、L3–L4=NDIS+TCP/IP、L7=WinSock |

### 第二段：Windows 驅動堆疊（Slides 7–14）

| # | 標題 | 內容重點 |
|---|------|---------|
| 7 | Windows 網路驅動架構總覽 | Application → WinSock → TDI/AFD → NDIS → Miniport Driver → NIC |
| 8 | OEM.sys 是什麼？ | 網卡廠商 Miniport Driver，直接與硬體 register 溝通 |
| 9 | OEM.sys 的資料路徑 | TX/RX 封包流向；CSI 原始資料在此層被韌體處理後上報 |
| 10 | nwifi.sys 是什麼？ | Windows Native WiFi 驅動，802.11 管理（認證、漫遊、省電） |
| 11 | nwifi.sys 的抽象化行為 | 向上只暴露 NDIS 標準介面，不轉發 PHY 層原始資訊（含 CSI） |
| 12 | NDIS 介面設計哲學 | NDIS OID 機制，跨廠商通用設計，代價是遮蔽底層細節 |
| 13 | WDI（WLAN Device Driver Interface） | Win10+ 新驅動模型（UMDF/KMDF 分離），CSI 支援依然缺席 |
| 14 | CSI 消失點分析 | OEM.sys ↔ nwifi.sys 邊界，無標準 OID 可查詢 CSI |

### 第三段：替代方案（Slides 15–22）

| # | 標題 | 內容重點 |
|---|------|---------|
| 15 | 為什麼 Linux 可以？ | cfg80211/mac80211、nl80211、部分驅動暴露 CSI debug interface |
| 16 | Linux CSI Tool（Intel 5300） | linux-80211n-csitool：修改韌體 + 自訂驅動，繞過 mac80211 直讀 CSI；基於 Ubuntu 10.04/kernel 2.6.36，歷史性工具 |
| 17 | 現代 CSI 工具生態（Linux-only） | PicoScenes（AX200/AX210，802.11ax，最完整）、FeitCSI（開源，所有格式/頻寬）、IAX（Intel AX200/201/210/211）、Nexmon CSI（Broadcom，802.11ac）；全部僅支援 Linux，強化 Windows 無法取得 CSI 的論點 |
| 18 | Windows 替代方案 1：Raw Packet Capture | WinPcap/Npcap + Monitor Mode，可得 802.11 frame header，無 CSI |
| 19 | Windows 替代方案 2：OEM 私有 IOCTL | 私有 DeviceIoControl 介面，需逆向 OEM.sys |
| 20 | Windows 替代方案 3：雙系統/VM | Linux 取 CSI，資料傳回 Windows 處理 |
| 21 | 替代方案比較表 | 難度 / 硬體需求 / 資料完整度 / 維護成本 |
| 22 | 硬體選型建議 | Intel IWL5300（linux-80211n-csitool）、Intel AX200/AX210（PicoScenes/FeitCSI）、Atheros AR9380（Atheros-CSI-Tool）、Broadcom BCM4339（Nexmon CSI） |

### 第四段：結語（Slides 23–25）

| # | 標題 | 內容重點 |
|---|------|---------|
| 23 | 架構限制總結 | 完整堆疊圖標注 CSI 消失點 |
| 24 | 給 RF 工程師的建議 | 選硬體先確認 CSI 支援、優先 Linux 環境、Windows 僅作後處理 |
| 25 | Q&A / 參考資料 | linux-80211n-csitool、Nexmon、Intel CSI Tool 連結 |

---

## 製作方式

**工具：** python-pptx（Python 程式化生成 .pptx）
**輸出位置：** `C:/Users/Jacky/Desktop/network-ppt-project/output/network-card-csi.pptx`
**設計風格：** 深色背景（深藍/黑），技術圖表為主，程式碼區塊使用等寬字型

### 關鍵視覺元素

- **Slide 6：** OSI 層次對應表（兩欄對比）
- **Slide 7：** Windows 驅動堆疊方塊圖（垂直層次）
- **Slide 14：** CSI 消失點標注圖（堆疊圖 + 紅色標記）
- **Slide 21：** 替代方案比較表（矩陣式）
- **Slide 23：** 完整架構總覽圖

---

## 技術內容說明

### CSI 消失原因（核心論點）

1. NIC 韌體計算 CSI 後儲存於硬體暫存器
2. OEM.sys 可透過 MMIO/PCI 讀取，但不上報給 nwifi.sys
3. nwifi.sys 只處理 802.11 管理幀，不轉發 PHY 層 metadata
4. NDIS 無標準 OID 定義 CSI 查詢（OID_802_11 系列不含 CSI）
5. WinSock/應用層完全無法感知 CSI 存在

### Linux 可行原因

- `iwlwifi` 驅動在 debug 模式下可將 CSI 寫入 debugfs
- `mac80211` 提供 radiotap header 擴充，部分實作包含 CSI 欄位
- 開源生態允許修改驅動 + 韌體

---

## 輸出規格

- 格式：`.pptx`
- 尺寸：16:9（1920×1080）
- 字型：標題 Calibri Bold 36pt，內文 Calibri 20pt，程式碼 Consolas 16pt
- 色彩：背景 #1a1a2e，主色 #4a9eff，強調 #ff6b6b，文字 #ffffff
