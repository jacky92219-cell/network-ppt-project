# src/theme.py
from pptx.util import Pt, Emu, Inches
from pptx.dml.color import RGBColor

# ─── 基礎色（純灰色系，R=G=B）────────────────────────────
BG_COLOR       = RGBColor(0xff, 0xff, 0xff)   # 純白背景
TITLE_BAR_BG   = RGBColor(0x2a, 0x2a, 0x2a)   # 純炭灰（v5.1 主色）
TEXT_COLOR     = RGBColor(0x1c, 0x1c, 0x1c)   # 近黑主文字
SUBTEXT_COLOR  = RGBColor(0x6b, 0x6b, 0x6b)   # 中灰副文字
ACCENT_COLOR   = RGBColor(0x50, 0x50, 0x50)   # 暗中灰 accent
DIVIDER_COLOR  = RGBColor(0xdc, 0xdc, 0xdc)   # 淺灰分隔線

# 面板色（內容背景）
PANEL_COLOR    = RGBColor(0xf5, 0xf5, 0xf5)   # 近白面板
PANEL_BORDER   = RGBColor(0xc8, 0xc8, 0xc8)   # 面板邊框灰

# 表格
TABLE_HEADER_BG = TITLE_BAR_BG
TABLE_ROW_ALT   = RGBColor(0xf0, 0xf0, 0xf0)  # 交替列灰

# 相容性別名
PRIMARY_COLOR  = ACCENT_COLOR
ACCENT2_COLOR  = ACCENT_COLOR
SOURCE_ANNOTATION_COLOR = RGBColor(0x80, 0x80, 0x80)  # 中灰（source 層標注）

# ─── 新增：白色常量 & 深色背景上的文字色 ────────────────
WHITE            = RGBColor(0xff, 0xff, 0xff)
TITLE_BAR_TEXT   = WHITE
FOOTER_TEXT      = RGBColor(0xb0, 0xb0, 0xb0)  # 標題列/footer 淺灰文字

# ─── 新增：Cover 裝飾色 ──────────────────────────────────
COVER_DECO_DARK    = RGBColor(0x1a, 0x1a, 0x1a)  # 右下裝飾方塊深灰
COVER_DECO_DARKER  = RGBColor(0x0f, 0x0f, 0x0f)  # 右下裝飾方塊極深灰
COVER_SUBTITLE     = RGBColor(0xc0, 0xc0, 0xc0)  # Cover 副標題淺灰
COVER_DATE         = RGBColor(0x90, 0x90, 0x90)  # Cover 日期中灰

# ─── 新增：漸層端點色（微妙亮度差）─────────────────────
TITLE_BAR_BG_LIGHT = RGBColor(0x36, 0x36, 0x36)  # 標題列右端（漸層終點）
COVER_BG_DARK      = RGBColor(0x1e, 0x1e, 0x1e)  # Cover 角落暗端
SECTION_BG_DARK    = RGBColor(0x22, 0x22, 0x22)  # Section break 頂端
SECTION_BG_LIGHT   = RGBColor(0x2e, 0x2e, 0x2e)  # Section break 底端
FOOTER_BG_LIGHT    = RGBColor(0x34, 0x34, 0x34)  # Footer 右端
PANEL_COLOR_DARK   = RGBColor(0xed, 0xed, 0xed)  # 面板底端（漸層終點）

# ─── 新增：Status 灰色系（用明度區分，不用色相）────────
STATUS_WARNING_FILL   = RGBColor(0x5a, 0x5a, 0x5a)  # Warning 填充（次暗）
STATUS_DANGER_FILL    = RGBColor(0x3a, 0x3a, 0x3a)  # Danger 填充（最暗=最嚴重）
STATUS_SOURCE_FILL    = RGBColor(0x70, 0x70, 0x70)  # Source 填充（最亮=資訊來源）
STATUS_WARNING_BORDER = RGBColor(0x88, 0x88, 0x88)  # Warning 邊框
STATUS_DANGER_BORDER  = RGBColor(0xaa, 0xaa, 0xaa)  # Danger 邊框（最亮=最高對比）
STATUS_SOURCE_BORDER  = RGBColor(0x80, 0x80, 0x80)  # Source 邊框
STATUS_DANGER_TEXT    = RGBColor(0x4a, 0x4a, 0x4a)  # Danger 標注文字
STATUS_WARNING_TEXT   = RGBColor(0x7a, 0x7a, 0x7a)  # Warning 標注文字

# ─── Section 識別色（統一灰色 accent）───────────────────
_UNIFIED_ACCENT = ACCENT_COLOR
SECTION_COLORS = {k: _UNIFIED_ACCENT for k in range(5)}

# ─── 標題列背景（統一炭灰）──────────────────────────────
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
