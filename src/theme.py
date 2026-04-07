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
SOURCE_ANNOTATION_COLOR = RGBColor(0x66, 0xff, 0x66)  # 綠色，用於堆疊圖 source 層標注

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
