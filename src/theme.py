# src/theme.py
from pptx.util import Pt, Emu, Inches
from pptx.dml.color import RGBColor

# 色彩
BG_COLOR       = RGBColor(0x1a, 0x1a, 0x2e)   # 深藍背景
PRIMARY_COLOR  = RGBColor(0x4a, 0x9e, 0xff)   # 主色（亮藍）
ACCENT_COLOR   = RGBColor(0xff, 0x6b, 0x6b)   # 強調（紅）
TEXT_COLOR     = RGBColor(0xff, 0xff, 0xff)   # 白色文字
SUBTEXT_COLOR  = RGBColor(0xaa, 0xcc, 0xff)   # 淡藍副文字
TABLE_HDR_BG   = RGBColor(0x0d, 0x47, 0xa1)   # 表頭深藍
TABLE_ROW_ALT  = RGBColor(0x1e, 0x2a, 0x4a)   # 交替列背景
SOURCE_ANNOTATION_COLOR = RGBColor(0x66, 0xff, 0x66)  # 綠色，用於堆疊圖 source 層標注

# 面板色（內容背景面板）
PANEL_COLOR    = RGBColor(0x10, 0x14, 0x25)   # 比 BG 更深的藍黑
PANEL_BORDER   = RGBColor(0x2a, 0x3a, 0x5c)   # 面板邊框色

# 段落識別色（Section Colors）
SECTION_COLORS = {
    1: RGBColor(0x29, 0xb6, 0xf6),  # 第一段：PHY/MAC 基礎 — 淡藍（區別於主色）
    2: RGBColor(0x9c, 0x6e, 0xff),  # 第二段：Windows 驅動堆疊 — 紫
    3: RGBColor(0x00, 0xd9, 0xaa),  # 第三段：替代方案 — 青綠
    4: RGBColor(0xff, 0xcc, 0x44),  # 第四段：結語 — 金
}
SECTION_BAR_WIDTH = Emu(91440)  # 左側段落色條寬度（0.1 inch）

# 標題列背景色（各段落的暗色變體）
TITLE_BAR_HEIGHT = Inches(1.15)
TITLE_BAR_COLORS = {
    0: RGBColor(0x10, 0x14, 0x25),   # 封面/通用
    1: RGBColor(0x0e, 0x3a, 0x52),   # 段落1：暗藍
    2: RGBColor(0x2a, 0x1a, 0x55),   # 段落2：暗紫
    3: RGBColor(0x00, 0x3a, 0x2e),   # 段落3：暗青綠
    4: RGBColor(0x3a, 0x2a, 0x00),   # 段落4：暗金
}

# 段落名稱
SECTION_NAMES = {
    0: "",
    1: "基礎架構",
    2: "Windows 驅動堆疊",
    3: "替代方案",
    4: "結語",
}

# 字型大小
TITLE_SIZE     = Pt(36)
SUBTITLE_SIZE  = Pt(24)
BODY_SIZE      = Pt(20)
SMALL_SIZE     = Pt(14)
CODE_SIZE      = Pt(16)

# 字型名稱
FONT_TITLE     = "Calibri"
FONT_BODY      = "Calibri"
FONT_CODE      = "Consolas"

# 投影片尺寸（16:9）
SLIDE_WIDTH    = Emu(9144000)   # 10 inches
SLIDE_HEIGHT   = Emu(5143500)   # 7.5 inches
