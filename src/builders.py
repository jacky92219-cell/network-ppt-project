# src/builders.py
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.oxml.ns import qn
from lxml import etree
import theme


def _add_arrowhead(connector):
    """在 connector 末端加箭頭"""
    sp = connector._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        return
    ln = spPr.find(qn('a:ln'))
    if ln is None:
        return
    for old in ln.findall(qn('a:tailEnd')):
        ln.remove(old)
    tailEnd = etree.SubElement(ln, qn('a:tailEnd'))
    tailEnd.set('type', 'arrow')
    tailEnd.set('w', 'med')
    tailEnd.set('len', 'med')

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

def add_title(slide, title_text, section: int = 0):
    """新增標準標題列"""
    left = Inches(0.4)
    top = Inches(0.15)
    width = Inches(9.2)
    height = Inches(1.0)  # 加高：呼吸空間
    add_textbox(slide, title_text, left, top, width, height,
                font_name=theme.FONT_TITLE, font_size=theme.TITLE_SIZE,
                color=get_title_color(section), bold=True)
    # 標題下方分隔線（配合加高後的位置）
    connector = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        left, Inches(1.1), left + width, Inches(1.1)
    )
    connector.line.color.rgb = get_title_color(section)
    connector.line.width = Pt(1.5)

def add_note(slide, note_text):
    """底部備註列"""
    left = Inches(0.4)
    top = Inches(6.8)
    width = Inches(9.2)
    height = Inches(0.5)
    add_textbox(slide, f"▶ {note_text}", left, top, width, height,
                font_size=theme.SMALL_SIZE, color=theme.SUBTEXT_COLOR)

def add_slide_number(slide, number: int, section: int = 0):
    """右下角頁碼"""
    sec_color = theme.SECTION_COLORS.get(section, theme.SUBTEXT_COLOR) if section > 0 else theme.SUBTEXT_COLOR
    add_textbox(slide, str(number),
                Inches(9.0), Inches(6.9), Inches(0.6), Inches(0.4),
                font_size=Pt(12), color=sec_color, align=PP_ALIGN.RIGHT)

def add_section_bar(slide, section: int):
    """左側段落識別色條"""
    if section == 0:
        return
    color = theme.SECTION_COLORS.get(section, theme.PRIMARY_COLOR)
    bar = slide.shapes.add_shape(1,
        Inches(0), Inches(0),
        theme.SECTION_BAR_WIDTH, theme.SLIDE_HEIGHT)
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()  # no border

def get_title_color(section: int) -> RGBColor:
    """根據段落取標題色"""
    if section == 0:
        return theme.PRIMARY_COLOR
    return theme.SECTION_COLORS.get(section, theme.PRIMARY_COLOR)

def build_cover(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_section_bar(slide, section)
    # 主標題：垂直居中偏上（2.0" 比原本 1.5" 更接近黃金比例）
    add_textbox(slide, data["title"],
                Inches(0.5), Inches(1.8), Inches(9.0), Inches(1.5),
                font_name=theme.FONT_TITLE, font_size=Pt(44),
                color=theme.PRIMARY_COLOR, bold=True, align=PP_ALIGN.CENTER)
    # 主副標之間的裝飾線
    dec_line = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(2.0), Inches(3.5), Inches(8.0), Inches(3.5)
    )
    dec_line.line.color.rgb = theme.SECTION_COLORS[1]
    dec_line.line.width = Pt(1.2)
    # 副標題：用 section 1 淡藍色
    add_textbox(slide, data["subtitle"],
                Inches(0.5), Inches(3.6), Inches(9.0), Inches(1.2),
                font_size=theme.SUBTITLE_SIZE,
                color=theme.SECTION_COLORS[1], align=PP_ALIGN.CENTER)
    add_textbox(slide, data["date"],
                Inches(0.5), Inches(6.5), Inches(9.0), Inches(0.5),
                font_size=theme.SMALL_SIZE,
                color=theme.SUBTEXT_COLOR, align=PP_ALIGN.CENTER)
    if data.get("_slide_num"):
        add_slide_number(slide, data["_slide_num"], section)

def build_bullets(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_section_bar(slide, section)
    add_title(slide, data["title"], section)
    left = Inches(0.5)
    top = Inches(1.3)   # 下移：呼吸空間
    width = Inches(9.0)
    height = Inches(5.3)
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
        if bullet == "":
            run.font.size = Pt(6)
            p.space_before = Pt(2)
        elif bullet.startswith("  "):
            run.font.size = Pt(15)
            run.font.color.rgb = theme.SUBTEXT_COLOR
            p.level = 1
            p.space_before = Pt(2)
        else:
            run.font.size = theme.BODY_SIZE
            run.font.color.rgb = theme.TEXT_COLOR
            p.space_before = Pt(5)
            if bullet.startswith("▶") or bullet.endswith("：") or (len(bullet) > 1 and bullet[0].isdigit()):
                run.font.bold = True
                run.font.color.rgb = theme.PRIMARY_COLOR
    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_slide_number(slide, data["_slide_num"], section)

def build_two_col(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_section_bar(slide, section)
    add_title(slide, data["title"], section)

    def add_col(title, bullets, left):
        add_textbox(slide, title, left, Inches(1.3), Inches(4.3), Inches(0.5),
                    font_size=Pt(18), color=theme.ACCENT_COLOR, bold=True)
        txBox = slide.shapes.add_textbox(left, Inches(1.9), Inches(4.3), Inches(4.6))
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
            if b == "":
                run.font.size = Pt(6)
                p.space_before = Pt(2)
            elif b.startswith("  "):
                run.font.size = Pt(14)
                run.font.color.rgb = theme.SUBTEXT_COLOR
                p.space_before = Pt(2)
            else:
                run.font.size = Pt(16)
                run.font.color.rgb = theme.TEXT_COLOR
                p.space_before = Pt(4)
                if b.startswith("✅") or b.startswith("❌") or b.startswith("⚠") or b.startswith("✗"):
                    run.font.bold = True

    add_col(data["left_title"], data["left_bullets"], Inches(0.3))
    div = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(4.8), Inches(1.1), Inches(4.8), Inches(6.6)
    )
    div.line.color.rgb = theme.PRIMARY_COLOR
    div.line.width = Pt(0.75)
    add_col(data["right_title"], data["right_bullets"], Inches(5.0))
    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_slide_number(slide, data["_slide_num"], section)

def build_table(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_section_bar(slide, section)
    add_title(slide, data["title"], section)
    rows = len(data["rows"]) + 1
    cols = len(data["headers"])
    left = Inches(0.3)
    top = Inches(1.3)
    width = Inches(9.4)
    height = Inches(0.5 * rows + 0.1)   # 加高：呼吸空間
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    col_width = int(width / cols)
    for i in range(cols):
        table.columns[i].width = col_width if i < cols - 1 else (width - col_width * (cols - 1))
    for ci, hdr in enumerate(data["headers"]):
        cell = table.cell(0, ci)
        cell.text = hdr
        cell.fill.solid()
        cell.fill.fore_color.rgb = theme.TABLE_HDR_BG
        cell.margin_left = Inches(0.08)
        cell.margin_right = Inches(0.08)
        cell.margin_top = Inches(0.05)
        cell.margin_bottom = Inches(0.05)
        p = cell.text_frame.paragraphs[0]
        run = p.runs[0] if p.runs else p.add_run()
        run.font.bold = True
        run.font.size = Pt(15)
        run.font.color.rgb = theme.TEXT_COLOR
        run.font.name = theme.FONT_BODY
        p.alignment = PP_ALIGN.CENTER
    for ri, row in enumerate(data["rows"]):
        bg = theme.BG_COLOR if ri % 2 == 0 else theme.TABLE_ROW_ALT
        for ci, val in enumerate(row):
            cell = table.cell(ri + 1, ci)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.04)
            p = cell.text_frame.paragraphs[0]
            run = p.runs[0] if p.runs else p.add_run()
            run.font.size = Pt(13)
            # 第一欄加粗並用主色，強化視覺層次
            if ci == 0:
                run.font.bold = True
                run.font.color.rgb = theme.PRIMARY_COLOR
            else:
                run.font.color.rgb = theme.TEXT_COLOR
            run.font.name = theme.FONT_BODY
    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_slide_number(slide, data["_slide_num"], section)

def build_flow(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_section_bar(slide, section)
    add_title(slide, data["title"], section)
    items = data["flow_items"]
    n = len(items)
    box_w = Inches(1.5)
    box_h = Inches(0.9)
    gap = Inches(0.2)
    total_w = n * box_w + (n - 1) * gap
    start_x = (theme.SLIDE_WIDTH - total_w) / 2
    y = Inches(2.5)
    for i, (label, desc) in enumerate(items):
        x = start_x + i * (box_w + gap)
        shape = slide.shapes.add_shape(1, x, y, box_w, box_h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = theme.TABLE_HDR_BG
        shape.line.color.rgb = theme.PRIMARY_COLOR
        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.06)
        tf.margin_right = Inches(0.06)
        tf.margin_top = Inches(0.08)
        tf.margin_bottom = Inches(0.08)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = label
        run.font.size = Pt(13)
        run.font.color.rgb = theme.TEXT_COLOR
        run.font.bold = True
        add_textbox(slide, desc, x, y + box_h + Inches(0.1),
                    box_w, Inches(0.4),
                    font_size=Pt(11), color=theme.SUBTEXT_COLOR,
                    align=PP_ALIGN.CENTER)
        if i < n - 1:
            # 用箭頭 Connector 取代填滿矩形
            arr_x1 = x + box_w
            arr_x2 = x + box_w + gap
            arr_y = y + box_h / 2
            arr = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT,
                arr_x1, arr_y, arr_x2, arr_y
            )
            arr.line.color.rgb = theme.PRIMARY_COLOR
            arr.line.width = Pt(1.5)
            _add_arrowhead(arr)
    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_slide_number(slide, data["_slide_num"], section)

def build_stack_diagram(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_section_bar(slide, section)
    add_title(slide, data["title"], section)
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
        label, sublabel, color_hex = layer_info
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
    if data.get("_slide_num"):
        add_slide_number(slide, data["_slide_num"], section)

def build_stack_diagram_annotated(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_section_bar(slide, section)
    add_title(slide, data["title"], section)
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
        ann_color = theme.ACCENT_COLOR if status in ("danger", "warning") else theme.SUBTEXT_COLOR
        if status == "source":
            ann_color = theme.SOURCE_ANNOTATION_COLOR
        add_textbox(slide, annotation,
                    Inches(6.0), y + Inches(0.1), Inches(3.8), box_h,
                    font_size=Pt(14), color=ann_color)
    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_slide_number(slide, data["_slide_num"], section)

def build_references(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_section_bar(slide, section)
    add_title(slide, data.get("qanda", "Q&A") + " / 參考資料", section)
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
    if data.get("_slide_num"):
        add_slide_number(slide, data["_slide_num"], section)

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
