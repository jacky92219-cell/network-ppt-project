# src/theme.py
from pptx.util import Pt, Emu, Inches
from pptx.dml.color import RGBColor

# ─── 基礎色 ──────────────────────────────────────────────
BG_COLOR       = RGBColor(0xff, 0xff, 0xff)   # 純白背景
TITLE_BAR_BG   = RGBColor(0x2c, 0x30, 0x38)   # 炭灰（v5.1 Charcoal）
TEXT_COLOR     = RGBColor(0x1a, 0x1a, 0x2e)   # 近黑主文字
SUBTEXT_COLOR  = RGBColor(0x5a, 0x6a, 0x7a)   # 中灰副文字
ACCENT_COLOR   = RGBColor(0x4a, 0x90, 0xd9)   # 天際藍 accent（v5.1 統一色）
DIVIDER_COLOR  = RGBColor(0xe0, 0xe5, 0xed)   # 淺灰分隔線

# 面板色（內容背景）
PANEL_COLOR    = RGBColor(0xf8, 0xfa, 0xfd)   # 近白淺藍
PANEL_BORDER   = RGBColor(0xd0, 0xd9, 0xe8)   # 淺藍灰邊框

# 表格
TABLE_HEADER_BG = TITLE_BAR_BG                 # 表頭同標題列
TABLE_ROW_ALT   = RGBColor(0xf4, 0xf7, 0xfb)  # 交替列極淺藍灰

# 保留相容性
PRIMARY_COLOR  = ACCENT_COLOR
ACCENT2_COLOR  = ACCENT_COLOR
SOURCE_ANNOTATION_COLOR = RGBColor(0x0e, 0x8f, 0x8f)  # 青藍（source 層標注）

# ─── Section 識別色 ──────────────────────────────────────
_UNIFIED_ACCENT = RGBColor(0x4a, 0x90, 0xd9)  # 統一 accent（天際藍，搭配炭灰標題列）
SECTION_COLORS = {
    0: _UNIFIED_ACCENT,
    1: _UNIFIED_ACCENT,
    2: _UNIFIED_ACCENT,
    3: _UNIFIED_ACCENT,
    4: _UNIFIED_ACCENT,
}

# ─── 標題列背景（v5 統一用 TITLE_BAR_BG）────────────────
TITLE_BAR_COLORS = {k: TITLE_BAR_BG for k in range(5)}

# ─── Section 名稱 ─────────────────────────────────────────
SECTION_NAMES = {
    0: "",
    1: "基礎架構",
    2: "Windows 驅動堆疊",
    3: "替代方案",
    4: "結語",
}

# ─── 字型 ────────────────────────────────────────────────
TITLE_SIZE    = Pt(28)
SUBTITLE_SIZE = Pt(24)
BODY_SIZE     = Pt(20)
SMALL_SIZE    = Pt(13)
CODE_SIZE     = Pt(16)
FONT_TITLE    = "Calibri"
FONT_BODY     = "Calibri"
FONT_CODE     = "Consolas"

# ─── 投影片尺寸（16:9）──────────────────────────────────
SLIDE_WIDTH    = Emu(9144000)   # 10 inches
SLIDE_HEIGHT   = Emu(5143500)   # 5.625 inches (16:9)

# ─── 集中式 Layout Grid ──────────────────────────────────
TITLE_BAR_H    = Inches(0.9)
MARGIN_H       = Inches(0.4)
CONTENT_LEFT   = MARGIN_H
CONTENT_RIGHT  = SLIDE_WIDTH - MARGIN_H
CONTENT_WIDTH  = CONTENT_RIGHT - CONTENT_LEFT
CONTENT_TOP    = TITLE_BAR_H + Inches(0.15)
FOOTER_H       = Inches(0.5)
CONTENT_BOTTOM = SLIDE_HEIGHT - FOOTER_H - Inches(0.04)
CONTENT_HEIGHT = CONTENT_BOTTOM - CONTENT_TOP
GUTTER         = Inches(0.35)
NOTE_HEIGHT    = Inches(0.28)
NOTE_TOP       = CONTENT_BOTTOM - NOTE_HEIGHT

# 相容性別名
SECTION_BAR_WIDTH = Emu(91440)
TITLE_BAR_HEIGHT  = TITLE_BAR_H
