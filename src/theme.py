# src/theme.py  — v5.1 Metal Gray（科技金屬灰）
from pptx.util import Pt, Emu, Inches
from pptx.dml.color import RGBColor

# ─── 基礎色：金屬深灰系 ──────────────────────────────────
BG_COLOR       = RGBColor(0xff, 0xff, 0xff)   # 內容區純白（保持可讀性）
TITLE_BAR_BG   = RGBColor(0x18, 0x18, 0x18)   # 近黑底（更深 = 更強科技感）
TEXT_COLOR     = RGBColor(0x1a, 0x1a, 0x1a)   # 內容區近黑文字
SUBTEXT_COLOR  = RGBColor(0x64, 0x64, 0x64)   # 中灰副文字
ACCENT_COLOR   = RGBColor(0x9a, 0x9a, 0x9a)   # 銀灰 accent（金屬感主色）
DIVIDER_COLOR  = RGBColor(0xd8, 0xd8, 0xd8)   # 淺灰分隔線

# 面板色
PANEL_COLOR    = RGBColor(0xfa, 0xfa, 0xfa)   # 近白面板
PANEL_BORDER   = RGBColor(0xb8, 0xb8, 0xb8)   # 邊框（比之前更深 = 更明顯）

# 表格
TABLE_HEADER_BG = TITLE_BAR_BG
TABLE_ROW_ALT   = RGBColor(0xf0, 0xf0, 0xf0)

# 相容性別名
PRIMARY_COLOR  = ACCENT_COLOR
ACCENT2_COLOR  = ACCENT_COLOR
SOURCE_ANNOTATION_COLOR = RGBColor(0x88, 0x88, 0x88)

# ─── 深色背景文字色 ───────────────────────────────────────
WHITE           = RGBColor(0xff, 0xff, 0xff)
TITLE_BAR_TEXT  = WHITE
FOOTER_TEXT     = RGBColor(0xc0, 0xc0, 0xc0)   # 銀白文字（比之前更亮 = 更清晰）

# ─── Cover 裝飾色 ──────────────────────────────────────────
COVER_DECO_DARK    = RGBColor(0x28, 0x28, 0x28)  # 右下深灰方塊
COVER_DECO_DARKER  = RGBColor(0x10, 0x10, 0x10)  # 右下極深方塊
COVER_SUBTITLE     = RGBColor(0xc8, 0xc8, 0xc8)  # 副標題銀白
COVER_DATE         = RGBColor(0x88, 0x88, 0x88)  # 日期中灰

# ─── 漸層端點色（金屬質感：更寬的明暗範圍）──────────────
TITLE_BAR_BG_LIGHT = RGBColor(0x42, 0x42, 0x42)  # 標題列右端（#18→#42，寬幅漸層）
COVER_BG_DARK      = RGBColor(0x0a, 0x0a, 0x0a)  # Cover 暗角（近黑）
SECTION_BG_DARK    = RGBColor(0x12, 0x12, 0x12)  # Section break 頂端
SECTION_BG_LIGHT   = RGBColor(0x28, 0x28, 0x28)  # Section break 底端
FOOTER_BG_LIGHT    = RGBColor(0x38, 0x38, 0x38)  # Footer 右端
PANEL_COLOR_DARK   = RGBColor(0xe8, 0xe8, 0xe8)  # 面板底端（加深漸層）

# ─── Status 灰色系（明度區分）──────────────────────────────
STATUS_WARNING_FILL   = RGBColor(0x55, 0x55, 0x55)
STATUS_DANGER_FILL    = RGBColor(0x30, 0x30, 0x30)   # 最暗（最嚴重）
STATUS_SOURCE_FILL    = RGBColor(0x72, 0x72, 0x72)   # 最亮（資訊來源）
STATUS_WARNING_BORDER = RGBColor(0x90, 0x90, 0x90)
STATUS_DANGER_BORDER  = RGBColor(0xb8, 0xb8, 0xb8)   # 最亮邊框（最高對比）
STATUS_SOURCE_BORDER  = RGBColor(0x82, 0x82, 0x82)
STATUS_DANGER_TEXT    = RGBColor(0x44, 0x44, 0x44)
STATUS_WARNING_TEXT   = RGBColor(0x78, 0x78, 0x78)

# ─── Section 識別色（銀灰 accent）───────────────────────────
_UNIFIED_ACCENT = ACCENT_COLOR
SECTION_COLORS  = {k: _UNIFIED_ACCENT for k in range(5)}
TITLE_BAR_COLORS = {k: TITLE_BAR_BG for k in range(5)}

# ─── Section 名稱 ──────────────────────────────────────────
SECTION_NAMES = {
    0: "",
    1: "基礎架構",
    2: "Windows 驅動堆疊",
    3: "替代方案",
    4: "結語",
}

# ─── 字型 ──────────────────────────────────────────────────
TITLE_SIZE    = Pt(28)
SUBTITLE_SIZE = Pt(24)
BODY_SIZE     = Pt(20)
SMALL_SIZE    = Pt(13)
CODE_SIZE     = Pt(16)
FONT_TITLE    = "Calibri"
FONT_BODY     = "Calibri"
FONT_CODE     = "Consolas"

# ─── 投影片尺寸（16:9）────────────────────────────────────
SLIDE_WIDTH    = Emu(9144000)   # 10 inches
SLIDE_HEIGHT   = Emu(5143500)   # 5.625 inches (16:9)

# ─── Layout Grid ───────────────────────────────────────────
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
