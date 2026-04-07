# src/builders.py
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
import theme


# ─────────────────────── XML helpers ───────────────────────

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
    tailEnd = OxmlElement('a:tailEnd')
    tailEnd.set('type', 'arrow')
    tailEnd.set('w', 'med')
    tailEnd.set('len', 'med')
    ln.append(tailEnd)


def _darken(color: RGBColor, factor: float = 0.6) -> RGBColor:
    """把 RGBColor 調暗（factor=0.6 → 60% 亮度）"""
    return RGBColor(
        int(color[0] * factor),
        int(color[1] * factor),
        int(color[2] * factor),
    )


# ─────────────────────── 基礎元件 ───────────────────────

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


# ─────────────────────── 共用裝飾元件（v2.0） ───────────────────────

def add_title_bar(slide, title_text: str, section: int = 0):
    """全寬標題列：暗色背景 + section 色底線 + 白色標題文字"""
    bar_h = theme.TITLE_BAR_HEIGHT
    bar_color = theme.TITLE_BAR_COLORS.get(section, theme.TITLE_BAR_COLORS[0])
    sec_color = theme.SECTION_COLORS.get(section, theme.PRIMARY_COLOR) if section > 0 else theme.PRIMARY_COLOR

    # 標題背景色塊
    bar = slide.shapes.add_shape(1,
        int(0), int(0),
        int(theme.SLIDE_WIDTH), int(bar_h))
    bar.fill.solid()
    bar.fill.fore_color.rgb = bar_color
    bar.line.fill.background()

    # 底部 section-color 亮線
    line_y = int(bar_h)
    connector = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        int(0), line_y,
        int(theme.SLIDE_WIDTH), line_y
    )
    connector.line.color.rgb = sec_color
    connector.line.width = Pt(2.0)

    # 標題文字
    add_textbox(slide, title_text,
                int(Inches(0.25)), int(Inches(0.12)),
                int(Inches(9.3)), int(bar_h - Inches(0.12)),
                font_name=theme.FONT_TITLE, font_size=theme.TITLE_SIZE,
                color=theme.TEXT_COLOR, bold=True)


def add_content_panel(slide, top, height, left=None, width=None):
    """繪製內容區域背景面板（深色圓角矩形）"""
    if left is None:
        left = int(Inches(0.25))
    if width is None:
        width = int(theme.SLIDE_WIDTH - Inches(0.5))
    panel = slide.shapes.add_shape(
        5,  # MSO_SHAPE_TYPE.ROUNDED_RECTANGLE
        int(left), int(top), int(width), int(height)
    )
    panel.fill.solid()
    panel.fill.fore_color.rgb = theme.PANEL_COLOR
    panel.line.color.rgb = theme.PANEL_BORDER
    panel.line.width = Pt(0.75)
    # 設定圓角調整
    adj = panel._element.spPr.find(qn('a:prstGeom'))
    if adj is not None:
        avLst = adj.find(qn('a:avLst'))
        if avLst is not None:
            for gd in avLst.findall(qn('a:gd')):
                avLst.remove(gd)
            gd_el = OxmlElement('a:gd')
            gd_el.set('name', 'adj')
            gd_el.set('fmla', 'val 16667')
            avLst.append(gd_el)
    return panel


def add_footer_bar(slide, number: int, section: int = 0):
    """底部資訊列：段落名稱 + 頁碼"""
    footer_h = int(Inches(0.35))
    footer_y = int(theme.SLIDE_HEIGHT - footer_h)
    sec_color = theme.SECTION_COLORS.get(section, theme.SUBTEXT_COLOR) if section > 0 else theme.PRIMARY_COLOR
    sec_name = theme.SECTION_NAMES.get(section, "")

    # 底部深色背景
    footer_bg = slide.shapes.add_shape(1,
        int(0), footer_y,
        int(theme.SLIDE_WIDTH), footer_h)
    footer_bg.fill.solid()
    footer_bg.fill.fore_color.rgb = theme.PANEL_COLOR
    footer_bg.line.fill.background()

    # 左側 section 色短線
    connector = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        int(0), footer_y,
        int(theme.SLIDE_WIDTH), footer_y
    )
    connector.line.color.rgb = sec_color
    connector.line.width = Pt(1.0)

    # 段落名稱
    if sec_name:
        add_textbox(slide, sec_name,
                    int(Inches(0.3)), footer_y,
                    int(Inches(5.0)), footer_h,
                    font_size=Pt(10), color=sec_color)

    # 頁碼
    add_textbox(slide, str(number),
                int(Inches(8.8)), footer_y,
                int(Inches(0.8)), footer_h,
                font_size=Pt(10), color=theme.SUBTEXT_COLOR,
                align=PP_ALIGN.RIGHT)


def add_note(slide, note_text):
    """底部備註列"""
    left = int(Inches(0.4))
    top = int(theme.SLIDE_HEIGHT - Inches(0.75))
    width = int(Inches(9.2))
    height = int(Inches(0.45))
    add_textbox(slide, f"▶ {note_text}", left, top, width, height,
                font_size=theme.SMALL_SIZE, color=theme.SUBTEXT_COLOR)


def add_section_bar(slide, section: int):
    """左側段落識別色條（保留相容性）"""
    if section == 0:
        return
    color = theme.SECTION_COLORS.get(section, theme.PRIMARY_COLOR)
    bar = slide.shapes.add_shape(1,
        int(0), int(0),
        int(theme.SECTION_BAR_WIDTH), int(theme.SLIDE_HEIGHT))
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()


def get_title_color(section: int) -> RGBColor:
    """根據段落取標題色"""
    if section == 0:
        return theme.PRIMARY_COLOR
    return theme.SECTION_COLORS.get(section, theme.PRIMARY_COLOR)


# ─────────────────────── Builders ───────────────────────

def build_cover(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)

    # 頂部裝飾幾何色塊群
    deco_colors = [
        theme.SECTION_COLORS[1],
        theme.SECTION_COLORS[2],
        theme.SECTION_COLORS[3],
        theme.SECTION_COLORS[4],
    ]
    deco_positions = [
        (int(Inches(0.1)), int(Inches(0.1)), int(Inches(1.4)), int(Inches(0.5))),
        (int(Inches(1.6)), int(Inches(0.2)), int(Inches(0.9)), int(Inches(0.35))),
        (int(Inches(0.3)), int(Inches(0.7)), int(Inches(0.6)), int(Inches(0.25))),
        (int(Inches(2.6)), int(Inches(0.1)), int(Inches(0.5)), int(Inches(0.45))),
    ]
    for i, (x, y, w, h) in enumerate(deco_positions):
        deco = slide.shapes.add_shape(1, x, y, w, h)
        deco.fill.solid()
        deco.fill.fore_color.rgb = _darken(deco_colors[i % len(deco_colors)], 0.5)
        deco.line.fill.background()

    # 主標題
    add_textbox(slide, data["title"],
                int(Inches(0.5)), int(Inches(1.6)), int(Inches(9.0)), int(Inches(1.8)),
                font_name=theme.FONT_TITLE, font_size=Pt(48),
                color=theme.PRIMARY_COLOR, bold=True, align=PP_ALIGN.CENTER)

    # accent 粗線
    dec_line = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        int(Inches(1.5)), int(Inches(3.5)),
        int(Inches(8.5)), int(Inches(3.5))
    )
    dec_line.line.color.rgb = theme.ACCENT_COLOR
    dec_line.line.width = Pt(3.0)

    # 副標題
    add_textbox(slide, data["subtitle"],
                int(Inches(0.5)), int(Inches(3.65)), int(Inches(9.0)), int(Inches(1.4)),
                font_size=theme.SUBTITLE_SIZE,
                color=theme.SECTION_COLORS[1], align=PP_ALIGN.CENTER)

    # 左右裝飾細線
    for x1, x2 in [(int(Inches(0.3)), int(Inches(2.5))), (int(Inches(6.5)), int(Inches(9.7)))]:
        line = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            x1, int(Inches(6.55)), x2, int(Inches(6.55))
        )
        line.line.color.rgb = theme.SUBTEXT_COLOR
        line.line.width = Pt(0.75)

    # 日期 + 版號
    date_ver = data["date"]
    if data.get("version"):
        date_ver = f"{data['date']}　　{data['version']}"
    add_textbox(slide, date_ver,
                int(Inches(0.5)), int(Inches(6.6)), int(Inches(9.0)), int(Inches(0.5)),
                font_size=theme.SMALL_SIZE,
                color=theme.SUBTEXT_COLOR, align=PP_ALIGN.CENTER)

    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


def build_section_break(slide, data):
    """段落過場投影片：全螢幕 section color 背景"""
    section = data.get("section", 1)
    sec_color = theme.SECTION_COLORS.get(section, theme.PRIMARY_COLOR)
    dark_color = _darken(sec_color, 0.25)

    # 全螢幕背景（section color 的暗色）
    set_slide_background(slide, dark_color)

    # 中央亮色大色塊
    panel_w = int(Inches(8.0))
    panel_h = int(Inches(2.4))
    panel_x = int((theme.SLIDE_WIDTH - panel_w) // 2)
    panel_y = int(Inches(1.8))
    panel = slide.shapes.add_shape(1, panel_x, panel_y, panel_w, panel_h)
    panel.fill.solid()
    panel.fill.fore_color.rgb = _darken(sec_color, 0.4)
    panel.line.color.rgb = sec_color
    panel.line.width = Pt(2.0)

    # 段落名稱（超大、居中）
    add_textbox(slide, data["title"],
                panel_x, panel_y,
                panel_w, int(Inches(1.5)),
                font_name=theme.FONT_TITLE, font_size=Pt(44),
                color=theme.TEXT_COLOR, bold=True, align=PP_ALIGN.CENTER)

    # 小字副標題
    add_textbox(slide, data.get("subtitle", ""),
                panel_x, panel_y + int(Inches(1.5)),
                panel_w, int(Inches(0.8)),
                font_size=Pt(18),
                color=sec_color, align=PP_ALIGN.CENTER)

    # 底部 section number 裝飾
    sec_num_text = f"Section {section}"
    add_textbox(slide, sec_num_text,
                int(Inches(0.3)), int(Inches(6.2)),
                int(Inches(2.0)), int(Inches(0.4)),
                font_size=Pt(12), color=sec_color)

    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


def build_bullets(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_title_bar(slide, data["title"], section)

    content_top = int(theme.TITLE_BAR_HEIGHT) + int(Inches(0.1))
    content_h = int(theme.SLIDE_HEIGHT - theme.TITLE_BAR_HEIGHT - Inches(0.6))
    add_content_panel(slide, content_top, content_h)

    left = int(Inches(0.5))
    top = int(theme.TITLE_BAR_HEIGHT) + int(Inches(0.2))
    width = int(Inches(8.8))
    height = content_h - int(Inches(0.15))

    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    sec_color = theme.SECTION_COLORS.get(section, theme.PRIMARY_COLOR) if section > 0 else theme.PRIMARY_COLOR

    for bullet in data["bullets"]:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        if bullet == "":
            run.text = bullet
            run.font.size = Pt(6)
            p.space_before = Pt(2)
        elif bullet.startswith("  "):
            run.text = "  › " + bullet.lstrip()
            run.font.size = Pt(15)
            run.font.color.rgb = theme.SUBTEXT_COLOR
            p.level = 1
            p.space_before = Pt(2)
        else:
            if bullet.startswith("▶") or bullet.endswith("：") or (len(bullet) > 1 and bullet[0].isdigit()):
                run.text = bullet
                run.font.bold = True
                run.font.color.rgb = sec_color
            else:
                run.text = "● " + bullet
                run.font.bold = False
                run.font.color.rgb = theme.TEXT_COLOR
            run.font.size = theme.BODY_SIZE
            p.space_before = Pt(5)

    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


def build_two_col(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_title_bar(slide, data["title"], section)

    content_top = int(theme.TITLE_BAR_HEIGHT) + int(Inches(0.1))
    content_h = int(theme.SLIDE_HEIGHT - theme.TITLE_BAR_HEIGHT - Inches(0.6))
    col_w = int(Inches(4.15))
    gap = int(Inches(0.5))
    left_x = int(Inches(0.25))
    right_x = left_x + col_w + gap

    # 左右各一個內容面板
    add_content_panel(slide, content_top, content_h, left=left_x, width=col_w)
    add_content_panel(slide, content_top, content_h, left=right_x, width=col_w)

    # 中間分隔線
    div_x = int(left_x + col_w + gap // 2)
    div = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        div_x, content_top,
        div_x, content_top + content_h
    )
    div.line.color.rgb = theme.PANEL_BORDER
    div.line.width = Pt(0.75)

    def add_col(title, bullets, col_left):
        sec_color = theme.SECTION_COLORS.get(section, theme.PRIMARY_COLOR) if section > 0 else theme.ACCENT_COLOR
        add_textbox(slide, title,
                    col_left + int(Inches(0.1)),
                    content_top + int(Inches(0.1)),
                    col_w - int(Inches(0.2)),
                    int(Inches(0.45)),
                    font_size=Pt(17), color=sec_color, bold=True)
        txBox = slide.shapes.add_textbox(
            col_left + int(Inches(0.1)),
            content_top + int(Inches(0.6)),
            col_w - int(Inches(0.2)),
            content_h - int(Inches(0.7))
        )
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
                run.font.size = Pt(13)
                run.font.color.rgb = theme.SUBTEXT_COLOR
                p.space_before = Pt(2)
            else:
                run.font.size = Pt(15)
                run.font.color.rgb = theme.TEXT_COLOR
                p.space_before = Pt(4)
                if b.startswith("✅") or b.startswith("❌") or b.startswith("⚠") or b.startswith("✗"):
                    run.font.bold = True

    add_col(data["left_title"], data["left_bullets"], left_x)
    add_col(data["right_title"], data["right_bullets"], right_x)

    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


def build_table(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_title_bar(slide, data["title"], section)

    sec_color = theme.SECTION_COLORS.get(section, theme.PRIMARY_COLOR) if section > 0 else theme.PRIMARY_COLOR
    rows = len(data["rows"]) + 1
    cols = len(data["headers"])
    left = int(Inches(0.25))
    top = int(theme.TITLE_BAR_HEIGHT) + int(Inches(0.15))
    width = int(Inches(9.5))
    height = int(Inches(0.52) * rows + Inches(0.1))
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    col_width = int(width / cols)
    for i in range(cols):
        table.columns[i].width = col_width if i < cols - 1 else (width - col_width * (cols - 1))

    for ci, hdr in enumerate(data["headers"]):
        cell = table.cell(0, ci)
        cell.text = hdr
        cell.fill.solid()
        cell.fill.fore_color.rgb = sec_color
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
            if ci == 0:
                run.font.bold = True
                run.font.color.rgb = sec_color
            else:
                run.font.color.rgb = theme.TEXT_COLOR
            run.font.name = theme.FONT_BODY

    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


def build_flow(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_title_bar(slide, data["title"], section)

    sec_color = theme.SECTION_COLORS.get(section, theme.PRIMARY_COLOR) if section > 0 else theme.PRIMARY_COLOR
    dark_sec = _darken(sec_color, 0.35)

    items = data["flow_items"]
    n = len(items)
    box_w = int(Inches(1.55))
    box_h = int(Inches(1.1))
    gap = int(Inches(0.18))
    total_w = n * box_w + (n - 1) * gap
    start_x = int((theme.SLIDE_WIDTH - total_w) // 2)
    y = int(Inches(2.2))

    for i, (label, desc) in enumerate(items):
        x = int(start_x + i * (box_w + gap))
        # 圓角矩形（shape type 5）
        shape = slide.shapes.add_shape(5, x, y, box_w, box_h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = dark_sec
        shape.line.color.rgb = sec_color
        shape.line.width = Pt(1.5)
        # 設定圓角
        adj = shape._element.spPr.find(qn('a:prstGeom'))
        if adj is not None:
            avLst = adj.find(qn('a:avLst'))
            if avLst is not None:
                for gd in avLst.findall(qn('a:gd')):
                    avLst.remove(gd)
                gd_el = OxmlElement('a:gd')
                gd_el.set('name', 'adj')
                gd_el.set('fmla', 'val 20000')
                avLst.append(gd_el)

        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = int(Inches(0.06))
        tf.margin_right = int(Inches(0.06))
        tf.margin_top = int(Inches(0.08))
        tf.margin_bottom = int(Inches(0.05))

        # label + desc 合在方塊內兩行
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = label
        run.font.size = Pt(12)
        run.font.color.rgb = theme.TEXT_COLOR
        run.font.bold = True

        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        run2 = p2.add_run()
        run2.text = desc
        run2.font.size = Pt(10)
        run2.font.color.rgb = sec_color

        if i < n - 1:
            arr_x1 = int(x + box_w)
            arr_x2 = int(x + box_w + gap)
            arr_y = int(y + box_h // 2)
            arr = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT,
                arr_x1, arr_y, arr_x2, arr_y
            )
            arr.line.color.rgb = sec_color
            arr.line.width = Pt(2.0)
            _add_arrowhead(arr)

    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


def build_stack_diagram(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_title_bar(slide, data["title"], section)

    layers = data["layers"]
    n = len(layers)
    box_h = int(Inches(0.55))
    gap = int(Inches(0.06))
    total_h = n * (box_h + gap)
    start_y = int((int(theme.SLIDE_HEIGHT) - total_h) // 2 + Inches(0.5))
    color_map = {
        "#2d5a27": RGBColor(0x2d, 0x5a, 0x27),
        "#1a3a6b": RGBColor(0x1a, 0x3a, 0x6b),
        "#4a2080": RGBColor(0x4a, 0x20, 0x80),
        "#7a1a1a": RGBColor(0x7a, 0x1a, 0x1a),
        "#3a3a00": RGBColor(0x3a, 0x3a, 0x00),
    }

    for i, layer_info in enumerate(layers):
        label, sublabel, color_hex = layer_info
        y = int(start_y + i * (box_h + gap))
        fill_color = color_map.get(color_hex, RGBColor(0x1a, 0x3a, 0x6b))
        border_color = theme.SECTION_COLORS.get(section, theme.PRIMARY_COLOR) if section > 0 else theme.PRIMARY_COLOR

        shape = slide.shapes.add_shape(1, int(Inches(1.0)), y, int(Inches(8.0)), box_h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
        shape.line.color.rgb = border_color
        shape.line.width = Pt(1.5)
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
                        int(Inches(9.2)), y, int(Inches(0.8)), box_h,
                        font_size=Pt(10), color=theme.SUBTEXT_COLOR)

        # 層間向下箭頭
        if i < n - 1:
            arr_x = int(Inches(4.95))
            arr_y1 = int(y + box_h)
            arr_y2 = int(y + box_h + gap)
            arr = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT,
                arr_x, arr_y1, arr_x, arr_y2
            )
            arr.line.color.rgb = border_color
            arr.line.width = Pt(1.5)

    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


def build_stack_diagram_annotated(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_title_bar(slide, data["title"], section)

    layers = data["layers"]
    n = len(layers)
    box_h = int(Inches(0.58))
    gap = int(Inches(0.06))
    start_y = int(theme.TITLE_BAR_HEIGHT) + int(Inches(0.15))
    status_colors = {
        "normal":  RGBColor(0x1a, 0x3a, 0x6b),
        "warning": RGBColor(0x6b, 0x4a, 0x00),
        "danger":  RGBColor(0x6b, 0x1a, 0x1a),
        "source":  RGBColor(0x1a, 0x5a, 0x27),
    }
    status_border = {
        "normal":  theme.PRIMARY_COLOR,
        "warning": RGBColor(0xff, 0xcc, 0x44),
        "danger":  theme.ACCENT_COLOR,
        "source":  theme.SOURCE_ANNOTATION_COLOR,
    }

    for i, (label, annotation, status) in enumerate(layers):
        y = int(start_y + i * (box_h + gap))
        fill_color = status_colors.get(status, status_colors["normal"])
        border_color = status_border.get(status, theme.PRIMARY_COLOR)
        border_w = Pt(2.5) if status in ("danger", "warning") else Pt(1.5)

        shape = slide.shapes.add_shape(1, int(Inches(0.3)), y, int(Inches(5.5)), box_h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
        shape.line.color.rgb = border_color
        shape.line.width = border_w
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
                    int(Inches(6.0)), y + int(Inches(0.1)), int(Inches(3.8)), box_h,
                    font_size=Pt(14), color=ann_color)

    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


def build_references(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_title_bar(slide, data.get("qanda", "Q&A") + " / 參考資料", section)

    content_top = int(theme.TITLE_BAR_HEIGHT) + int(Inches(0.15))
    content_h = int(theme.SLIDE_HEIGHT - theme.TITLE_BAR_HEIGHT - Inches(0.5))
    add_content_panel(slide, content_top, content_h)

    txBox = slide.shapes.add_textbox(
        int(Inches(0.5)),
        content_top + int(Inches(0.15)),
        int(Inches(9.0)),
        content_h - int(Inches(0.3))
    )
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
        p.space_before = Pt(3)

    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


BUILDERS = {
    "cover":                    build_cover,
    "section_break":            build_section_break,
    "bullets":                  build_bullets,
    "two_col":                  build_two_col,
    "table":                    build_table,
    "flow":                     build_flow,
    "stack_diagram":            build_stack_diagram,
    "stack_diagram_annotated":  build_stack_diagram_annotated,
    "references":               build_references,
}
