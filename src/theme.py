# src/theme.py
from pptx.util import Pt, Emu, Inches
from pptx.dml.color import RGBColor

# ─── 色彩 ───────────────────────────────────────────────────
BG_COLOR       = RGBColor(0x1a, 0x1a, 0x2e)   # 深藍背景
PRIMARY_COLOR  = RGBColor(0x00, 0xb4, 0xff)   # 電光藍（升級）
ACCENT_COLOR   = RGBColor(0xff, 0x47, 0x57)   # 鮮紅（升級）
ACCENT2_COLOR  = RGBColor(0x00, 0xff, 0xc8)   # 螢光青綠（新增）
TEXT_COLOR     = RGBColor(0xf0, 0xf0, 0xf0)   # 微降亮白
SUBTEXT_COLOR  = RGBColor(0x8a, 0xb4, 0xf8)   # 柔和藍副文字
TABLE_ROW_ALT  = RGBColor(0x1e, 0x2a, 0x4a)   # 交替列背景
SOURCE_ANNOTATION_COLOR = RGBColor(0x66, 0xff, 0x66)  # 綠色，堆疊圖 source 層標注

# 面板色
PANEL_COLOR    = RGBColor(0x0d, 0x11, 0x17)   # GitHub Dark 深色面板
PANEL_BORDER   = RGBColor(0x30, 0x36, 0x3d)   # GitHub Dark 邊框

# 段落識別色（更飽和）
SECTION_COLORS = {
    1: RGBColor(0x00, 0xb4, 0xff),  # Section 1：電光藍
    2: RGBColor(0xa8, 0x55, 0xf7),  # Section 2：鮮紫
    3: RGBColor(0x00, 0xff, 0xc8),  # Section 3：螢光青綠
    4: RGBColor(0xfb, 0xbf, 0x24),  # Section 4：琥珀金
}

# 標題列背景色（各段落暗色變體）
TITLE_BAR_COLORS = {
    0: RGBColor(0x0d, 0x11, 0x17),
    1: RGBColor(0x00, 0x2a, 0x44),   # 電光藍暗版
    2: RGBColor(0x1e, 0x0a, 0x33),   # 鮮紫暗版
    3: RGBColor(0x00, 0x2e, 0x22),   # 青綠暗版
    4: RGBColor(0x2a, 0x1e, 0x00),   # 琥珀金暗版
}

# 段落名稱
SECTION_NAMES = {
    0: "",
    1: "基礎架構",
    2: "Windows 驅動堆疊",
    3: "替代方案",
    4: "結語",
}

# ─── 字型 ───────────────────────────────────────────────────
TITLE_SIZE     = Pt(36)
SUBTITLE_SIZE  = Pt(24)
BODY_SIZE      = Pt(20)
SMALL_SIZE     = Pt(14)
CODE_SIZE      = Pt(16)
FONT_TITLE     = "Calibri"
FONT_BODY      = "Calibri"
FONT_CODE      = "Consolas"

# ─── 投影片尺寸（16:9）───────────────────────────────────────
SLIDE_WIDTH    = Emu(9144000)   # 10 inches
SLIDE_HEIGHT   = Emu(5143500)   # 5.625 inches (16:9)  ← 修正：原註解誤寫為 7.5"

# ─── 集中式 Layout Grid ─────────────────────────────────────
TITLE_BAR_H    = Inches(1.0)                              # 標題列高度
MARGIN_H       = Inches(0.4)                              # 左右對稱 margin
CONTENT_LEFT   = MARGIN_H                                 # 0.4"
CONTENT_RIGHT  = SLIDE_WIDTH - MARGIN_H                   # 9.6"
CONTENT_WIDTH  = CONTENT_RIGHT - CONTENT_LEFT             # 9.2"
CONTENT_TOP    = TITLE_BAR_H + Inches(0.12)               # 1.12"
FOOTER_H       = Inches(0.585)                            # 底部資訊列高度
CONTENT_BOTTOM = SLIDE_HEIGHT - FOOTER_H - Inches(0.05)   # 4.99"
CONTENT_HEIGHT = CONTENT_BOTTOM - CONTENT_TOP             # 3.87"
GUTTER         = Inches(0.35)                             # 欄間距
NOTE_HEIGHT    = Inches(0.30)                             # 備註文字高度
NOTE_TOP       = CONTENT_BOTTOM - NOTE_HEIGHT             # 4.69"

# 相容性（舊名稱，部分 builder 過渡期用）
SECTION_BAR_WIDTH = Emu(91440)  # 0.1 inch
TITLE_BAR_HEIGHT  = TITLE_BAR_H  # 舊別名
