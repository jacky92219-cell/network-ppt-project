# src/builders.py
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
import theme


# ─────────────────────── XML helpers ───────────────────────

def _add_arrowhead(connector):
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


def _darken(color: RGBColor, factor: float = 0.5) -> RGBColor:
    return RGBColor(
        int(color[0] * factor),
        int(color[1] * factor),
        int(color[2] * factor),
    )


def _lighten(color: RGBColor, factor: float = 1.5) -> RGBColor:
    return RGBColor(
        min(255, int(color[0] * factor)),
        min(255, int(color[1] * factor)),
        min(255, int(color[2] * factor)),
    )


def _set_gradient_fill(shape, color1: RGBColor, color2: RGBColor, angle: float = 0):
    """微妙漸層填充。angle: 0=左→右, 270=上→下, 225=左上→右下"""
    fill = shape.fill
    fill.gradient()
    fill.gradient_stops[0].color.rgb = color1
    fill.gradient_stops[0].position = 0.0
    fill.gradient_stops[1].color.rgb = color2
    fill.gradient_stops[1].position = 1.0
    fill.gradient_angle = angle


def _set_rounded_corner(shape, val: int = 16667):
    adj = shape._element.spPr.find(qn('a:prstGeom'))
    if adj is None:
        return
    avLst = adj.find(qn('a:avLst'))
    if avLst is None:
        return
    for gd in avLst.findall(qn('a:gd')):
        avLst.remove(gd)
    gd_el = OxmlElement('a:gd')
    gd_el.set('name', 'adj')
    gd_el.set('fmla', f'val {val}')
    avLst.append(gd_el)


# ─────────────────────── 基礎元件 ───────────────────────

def set_slide_background(slide, color: RGBColor):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_slide_image(slide, image_path: str, left, top, width, height):
    """將圖片加入 slide 指定位置（維持原始比例）"""
    slide.shapes.add_picture(image_path, int(left), int(top), int(width), int(height))


def add_textbox(slide, text, left, top, width, height,
                font_name=theme.FONT_BODY, font_size=theme.BODY_SIZE,
                color=theme.TEXT_COLOR, bold=False, align=PP_ALIGN.LEFT,
                word_wrap=True):
    txBox = slide.shapes.add_textbox(
        int(left), int(top), int(width), int(height)
    )
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


# ─────────────────────── Grid-aware 共用裝飾元件 ───────────────────────

def add_title_bar(slide, title_text: str, section: int = 0):
    bar_h = int(theme.TITLE_BAR_H)
    sec_color = theme.SECTION_COLORS.get(section, theme.ACCENT_COLOR)

    bar = slide.shapes.add_shape(1, 0, 0, int(theme.SLIDE_WIDTH), bar_h)
    _set_gradient_fill(bar, theme.TITLE_BAR_BG, theme.TITLE_BAR_BG_LIGHT, angle=0)
    bar.line.fill.background()

    # 頂部銀白高光線（1px，金屬感）
    highlight = slide.shapes.add_shape(1, 0, 0, int(theme.SLIDE_WIDTH), int(Inches(0.012)))
    _set_gradient_fill(highlight,
                       RGBColor(0x70, 0x70, 0x70),
                       RGBColor(0x30, 0x30, 0x30), angle=0)
    highlight.line.fill.background()

    accent_w = int(Inches(0.055))
    acc = slide.shapes.add_shape(1, 0, 0, accent_w, bar_h)
    _set_gradient_fill(acc, theme.ACCENT_COLOR, theme.TITLE_BAR_BG_LIGHT, angle=270)
    acc.line.fill.background()

    add_textbox(slide, title_text,
                int(theme.CONTENT_LEFT),
                int(Inches(0.12)),
                int(theme.CONTENT_WIDTH - Inches(1.5)),
                bar_h - int(Inches(0.1)),
                font_name=theme.FONT_TITLE, font_size=theme.TITLE_SIZE,
                color=theme.TITLE_BAR_TEXT, bold=True)

    if section > 0:
        sec_label = f"Section {section}/4"
        add_textbox(slide, sec_label,
                    int(theme.CONTENT_RIGHT - Inches(1.4)),
                    int(Inches(0.12)),
                    int(Inches(1.4)),
                    bar_h - int(Inches(0.1)),
                    font_size=theme.SMALL_SIZE,
                    color=theme.FOOTER_TEXT,
                    align=PP_ALIGN.RIGHT)


def add_content_panel(slide, top=None, height=None, left=None, width=None):
    if left is None:
        left = int(theme.CONTENT_LEFT)
    if width is None:
        width = int(theme.CONTENT_WIDTH)
    if top is None:
        top = int(theme.CONTENT_TOP)
    if height is None:
        height = int(theme.CONTENT_HEIGHT)

    panel = slide.shapes.add_shape(5, int(left), int(top), int(width), int(height))
    _set_gradient_fill(panel, theme.PANEL_COLOR, theme.PANEL_COLOR_DARK, angle=270)
    panel.line.color.rgb = theme.PANEL_BORDER
    panel.line.width = Pt(0.75)
    _set_rounded_corner(panel, 8000)
    return panel


def add_footer_bar(slide, number: int, section: int = 0):
    footer_h = int(theme.FOOTER_H)
    footer_y = int(theme.SLIDE_HEIGHT - theme.FOOTER_H)
    sec_name = theme.SECTION_NAMES.get(section, "")

    footer_bg = slide.shapes.add_shape(1, 0, footer_y, int(theme.SLIDE_WIDTH), footer_h)
    _set_gradient_fill(footer_bg, theme.TITLE_BAR_BG, theme.FOOTER_BG_LIGHT, angle=0)
    footer_bg.line.fill.background()

    text_h = int(footer_h - Inches(0.08))
    text_y = footer_y + int(Inches(0.04))

    if sec_name:
        add_textbox(slide, sec_name,
                    int(Inches(0.3)), text_y,
                    int(Inches(5.0)), text_h,
                    font_size=theme.SMALL_SIZE, color=theme.FOOTER_TEXT)

    add_textbox(slide, f"{number:02d} / 25",
                int(theme.CONTENT_RIGHT - Inches(1.0)),
                text_y,
                int(Inches(1.0)), text_h,
                font_size=theme.SMALL_SIZE, color=theme.FOOTER_TEXT,
                align=PP_ALIGN.RIGHT)


def add_note(slide, note_text):
    line_y = int(theme.NOTE_TOP) - 4
    div = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, int(theme.CONTENT_LEFT), line_y,
        int(theme.CONTENT_RIGHT), line_y)
    div.line.color.rgb = theme.DIVIDER_COLOR
    div.line.width = Pt(0.5)

    add_textbox(slide, note_text,
                int(theme.CONTENT_LEFT + Inches(0.1)),
                int(theme.NOTE_TOP),
                int(theme.CONTENT_WIDTH - Inches(0.2)),
                int(theme.NOTE_HEIGHT),
                font_size=theme.SMALL_SIZE, color=theme.SUBTEXT_COLOR)


# ─────────────────────── Builders ───────────────────────

def build_cover(slide, data):
    section = data.get("section", 0)
    sw = int(theme.SLIDE_WIDTH)
    sh = int(theme.SLIDE_HEIGHT)
    sec_color = theme.SECTION_COLORS.get(1, theme.ACCENT_COLOR)

    # 全版漸層背景（斜角：左上暗→右下稍亮）
    bg = slide.shapes.add_shape(1, 0, 0, sw, sh)
    _set_gradient_fill(bg, theme.COVER_BG_DARK, theme.TITLE_BAR_BG, angle=225)
    bg.line.fill.background()

    # 右下金屬裝飾方塊（漸層）
    deco1 = slide.shapes.add_shape(1,
        int(sw - Inches(3.0)), int(sh - Inches(2.5)),
        int(Inches(3.0)), int(Inches(2.5)))
    _set_gradient_fill(deco1, theme.COVER_DECO_DARK, theme.COVER_DECO_DARKER, angle=225)
    deco1.line.fill.background()

    deco2 = slide.shapes.add_shape(1,
        int(sw - Inches(1.8)), int(sh - Inches(1.5)),
        int(Inches(1.8)), int(Inches(1.5)))
    _set_gradient_fill(deco2, theme.COVER_DECO_DARKER, RGBColor(0x05, 0x05, 0x05), angle=225)
    deco2.line.fill.background()

    # 左側金屬 accent 條（上亮下暗）
    top_acc = slide.shapes.add_shape(1, 0, 0, int(Inches(0.08)), sh)
    _set_gradient_fill(top_acc, theme.ACCENT_COLOR, theme.TITLE_BAR_BG, angle=270)
    top_acc.line.fill.background()

    # 水平金屬高光線（Cover 專屬）
    shine = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, 0, int(Inches(1.0)), sw, int(Inches(1.0)))
    shine.line.color.rgb = RGBColor(0x40, 0x40, 0x40)
    shine.line.width = Pt(0.5)

    add_textbox(slide, data["title"],
                int(Inches(0.5)), int(Inches(1.4)),
                int(Inches(9.0)), int(Inches(1.5)),
                font_name=theme.FONT_TITLE, font_size=Pt(40),
                color=theme.TITLE_BAR_TEXT, bold=True,
                align=PP_ALIGN.LEFT)

    line_y = int(Inches(3.05))
    h_line = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, int(Inches(0.5)), line_y,
        int(Inches(9.5)), line_y)
    h_line.line.color.rgb = sec_color
    h_line.line.width = Pt(2.0)

    add_textbox(slide, data["subtitle"],
                int(Inches(0.5)), int(Inches(3.15)),
                int(Inches(9.0)), int(Inches(1.2)),
                font_size=Pt(22),
                color=theme.COVER_SUBTITLE,
                align=PP_ALIGN.LEFT)

    date_ver = data["date"]
    if data.get("version"):
        date_ver = f"{data['date']}    {data['version']}"
    add_textbox(slide, date_ver,
                int(Inches(0.5)), int(Inches(4.45)),
                int(Inches(5.0)), int(Inches(0.45)),
                font_size=Pt(14),
                color=theme.COVER_DATE)

    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


def build_section_break(slide, data):
    section = data.get("section", 1)
    sec_color = theme.SECTION_COLORS.get(section, theme.ACCENT_COLOR)
    sw = int(theme.SLIDE_WIDTH)
    sh = int(theme.SLIDE_HEIGHT)

    # 全版漸層背景（上暗→下稍亮）
    bg = slide.shapes.add_shape(1, 0, 0, sw, sh)
    _set_gradient_fill(bg, theme.SECTION_BG_DARK, theme.SECTION_BG_LIGHT, angle=270)
    bg.line.fill.background()

    # 頂部金屬漸層條
    top_bar = slide.shapes.add_shape(1, 0, 0, sw, int(Inches(0.12)))
    _set_gradient_fill(top_bar, theme.ACCENT_COLOR, RGBColor(0x50, 0x50, 0x50), angle=0)
    top_bar.line.fill.background()

    # 銀色高光細線
    shine = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, 0, int(Inches(0.12)), sw, int(Inches(0.12)))
    shine.line.color.rgb = RGBColor(0x60, 0x60, 0x60)
    shine.line.width = Pt(0.5)

    num_color = RGBColor(
        min(255, int(sec_color[0] * 0.25) + int(theme.TITLE_BAR_BG[0] * 0.75)),
        min(255, int(sec_color[1] * 0.25) + int(theme.TITLE_BAR_BG[1] * 0.75)),
        min(255, int(sec_color[2] * 0.25) + int(theme.TITLE_BAR_BG[2] * 0.75)),
    )
    add_textbox(slide, str(section),
                int(Inches(6.5)), int(Inches(0.8)),
                int(Inches(3.0)), int(Inches(3.5)),
                font_name=theme.FONT_TITLE, font_size=Pt(160),
                color=num_color, bold=True, align=PP_ALIGN.RIGHT)

    add_textbox(slide, data["title"],
                int(Inches(0.6)), int(Inches(1.8)),
                int(Inches(7.0)), int(Inches(1.3)),
                font_name=theme.FONT_TITLE, font_size=Pt(44),
                color=theme.TITLE_BAR_TEXT, bold=True)

    add_textbox(slide, data.get("subtitle", ""),
                int(Inches(0.6)), int(Inches(3.1)),
                int(Inches(7.5)), int(Inches(0.9)),
                font_size=Pt(22),
                color=sec_color)

    bottom_y = int(Inches(4.2))
    b_line = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, int(Inches(0.4)), bottom_y,
        int(Inches(9.6)), bottom_y)
    b_line.line.color.rgb = sec_color
    b_line.line.width = Pt(1.5)

    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


def build_bullets(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_title_bar(slide, data["title"], section)

    sec_color = theme.SECTION_COLORS.get(section, theme.ACCENT_COLOR)

    stripe_x = int(theme.CONTENT_LEFT)
    stripe_w = int(Inches(0.04))
    stripe = slide.shapes.add_shape(1,
        stripe_x, int(theme.CONTENT_TOP),
        stripe_w, int(theme.NOTE_TOP - theme.CONTENT_TOP))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = sec_color
    stripe.line.fill.background()

    txBox = slide.shapes.add_textbox(
        int(theme.CONTENT_LEFT + Inches(0.18)),
        int(theme.CONTENT_TOP + Inches(0.1)),
        int(theme.CONTENT_WIDTH - Inches(0.3)),
        int(theme.NOTE_TOP - theme.CONTENT_TOP - Inches(0.2))
    )
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
        if bullet == "":
            run.text = bullet
            run.font.size = Pt(6)
        elif bullet.startswith("  "):
            run.text = "  · " + bullet.lstrip()
            run.font.size = Pt(15)
            run.font.color.rgb = theme.SUBTEXT_COLOR
            p.level = 1
            p.space_before = Pt(1)
        else:
            if (bullet.startswith("▶") or bullet.startswith("⚠")
                    or (len(bullet) > 1 and bullet[0].isdigit())):
                run.text = bullet
                run.font.bold = True
                run.font.color.rgb = sec_color
            elif bullet.startswith("❌") or bullet.startswith("✅"):
                run.text = bullet
                run.font.bold = True
                run.font.color.rgb = theme.TEXT_COLOR
            else:
                run.text = "● " + bullet.lstrip("● ")
                run.font.color.rgb = theme.TEXT_COLOR
            run.font.size = Pt(17)
            p.space_before = Pt(4)

    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


def build_two_col(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_title_bar(slide, data["title"], section)

    sec_color = theme.SECTION_COLORS.get(section, theme.ACCENT_COLOR)
    col_w = int((theme.CONTENT_WIDTH - theme.GUTTER) // 2)
    left_x = int(theme.CONTENT_LEFT)
    right_x = int(theme.CONTENT_LEFT + col_w + theme.GUTTER)
    ct = int(theme.CONTENT_TOP)
    note_top = int(theme.NOTE_TOP)
    ch = note_top - ct - int(Inches(0.05))

    div_x = int(left_x + col_w + theme.GUTTER // 2)
    div = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, div_x, ct, div_x, note_top)
    div.line.color.rgb = theme.DIVIDER_COLOR
    div.line.width = Pt(0.5)

    def add_col(title, bullets, col_left, title_color):
        add_textbox(slide, title,
                    col_left + int(Inches(0.1)),
                    ct + int(Inches(0.08)),
                    col_w - int(Inches(0.2)),
                    int(Inches(0.42)),
                    font_size=Pt(18), color=title_color, bold=True)
        txBox = slide.shapes.add_textbox(
            col_left + int(Inches(0.1)),
            ct + int(Inches(0.55)),
            col_w - int(Inches(0.2)),
            ch - int(Inches(0.6))
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
            elif b.startswith("  "):
                run.font.size = Pt(13)
                run.font.color.rgb = theme.SUBTEXT_COLOR
                p.space_before = Pt(1)
            else:
                run.font.size = Pt(15)
                run.font.color.rgb = theme.TEXT_COLOR
                p.space_before = Pt(3)
                if b.startswith(("✅", "❌", "⚠", "✗")):
                    run.font.bold = True

    add_col(data["left_title"], data["left_bullets"], left_x, sec_color)
    add_col(data["right_title"], data["right_bullets"], right_x, theme.ACCENT_COLOR)

    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


def build_table(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_title_bar(slide, data["title"], section)

    sec_color = theme.SECTION_COLORS.get(section, theme.ACCENT_COLOR)
    rows = len(data["rows"]) + 1
    cols = len(data["headers"])
    left = int(theme.CONTENT_LEFT + Inches(0.05))
    top = int(theme.CONTENT_TOP + Inches(0.08))
    width = int(theme.CONTENT_WIDTH - Inches(0.1))
    note_reserve = int(theme.NOTE_HEIGHT + Inches(0.1)) if "note" in data else 0
    max_h = int(theme.CONTENT_HEIGHT - Inches(0.1)) - note_reserve
    height = min(int(Inches(0.48) * rows + Inches(0.1)), max_h)

    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    col_width = int(width / cols)
    for i in range(cols):
        table.columns[i].width = col_width if i < cols - 1 else (width - col_width * (cols - 1))

    for ci, hdr in enumerate(data["headers"]):
        cell = table.cell(0, ci)
        cell.text = hdr
        cell.fill.solid()
        cell.fill.fore_color.rgb = theme.TABLE_HEADER_BG
        cell.margin_left = Inches(0.08)
        cell.margin_right = Inches(0.08)
        cell.margin_top = Inches(0.05)
        cell.margin_bottom = Inches(0.05)
        p = cell.text_frame.paragraphs[0]
        run = p.runs[0] if p.runs else p.add_run()
        run.font.bold = True
        run.font.size = Pt(15)
        run.font.color.rgb = theme.WHITE
        run.font.name = theme.FONT_BODY
        p.alignment = PP_ALIGN.CENTER

    for ri, row in enumerate(data["rows"]):
        bg = theme.WHITE if ri % 2 == 0 else theme.TABLE_ROW_ALT
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
            run.font.size = Pt(14)
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

    sec_color = theme.SECTION_COLORS.get(section, theme.ACCENT_COLOR)

    items = data["flow_items"]
    n = len(items)
    box_w = int(Inches(1.5))
    gap = int(Inches(0.35))
    box_h = int(Inches(1.1))
    total_w = n * box_w + (n - 1) * gap
    start_x = int((theme.SLIDE_WIDTH - total_w) // 2)
    note_reserve = int(theme.NOTE_HEIGHT + Inches(0.15)) if "note" in data else 0
    avail_h = int(theme.CONTENT_HEIGHT) - note_reserve
    y = int(theme.CONTENT_TOP + (avail_h - box_h) // 2)

    for i, (label, desc) in enumerate(items):
        x = int(start_x + i * (box_w + gap))

        shape = slide.shapes.add_shape(5, x, y, box_w, box_h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = theme.WHITE
        shape.line.color.rgb = sec_color
        shape.line.width = Pt(1.5)
        _set_rounded_corner(shape, 20000)

        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = int(Inches(0.06))
        tf.margin_right = int(Inches(0.06))
        tf.margin_top = int(Inches(0.1))
        tf.margin_bottom = int(Inches(0.05))

        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = label
        run.font.size = Pt(13)
        run.font.color.rgb = theme.TEXT_COLOR
        run.font.bold = True
        run.font.name = theme.FONT_BODY

        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        run2 = p2.add_run()
        run2.text = desc
        run2.font.size = Pt(11)
        run2.font.color.rgb = sec_color
        run2.font.name = theme.FONT_BODY

        if i < n - 1:
            arr_x1 = int(x + box_w)
            arr_x2 = int(x + box_w + gap)
            arr_y = int(y + box_h // 2)
            arr = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT, arr_x1, arr_y, arr_x2, arr_y)
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

    sec_color = theme.SECTION_COLORS.get(section, theme.ACCENT_COLOR)
    layers = data["layers"]
    n = len(layers)
    box_h = int(Inches(0.42))
    gap = int(Inches(0.04))
    total_h = n * (box_h + gap) - gap
    note_reserve = int(theme.NOTE_HEIGHT + Inches(0.1)) if "note" in data else 0
    avail_h = int(theme.CONTENT_HEIGHT) - note_reserve
    start_y = int(theme.CONTENT_TOP + (avail_h - total_h) // 2)

    base = sec_color
    layer_colors = []
    for i in range(n):
        factor = 0.35 + (i / max(n - 1, 1)) * 0.55
        layer_colors.append(RGBColor(
            min(255, int(base[0] * factor)),
            min(255, int(base[1] * factor)),
            min(255, int(base[2] * factor)),
        ))

    box_left = int(theme.CONTENT_LEFT + Inches(0.3))
    box_w = int(Inches(7.2))

    for i, layer_info in enumerate(layers):
        label, sublabel, _ = layer_info
        y = int(start_y + i * (box_h + gap))
        fill_color = layer_colors[i]

        shape = slide.shapes.add_shape(1, box_left, y, box_w, box_h)
        _set_gradient_fill(shape, fill_color, _lighten(fill_color, 1.15), angle=0)
        shape.line.color.rgb = sec_color
        shape.line.width = Pt(1.0)

        tf = shape.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = label
        run.font.size = Pt(14)
        run.font.color.rgb = theme.WHITE
        run.font.bold = True
        run.font.name = theme.FONT_BODY

        if sublabel:
            sub_left = int(box_left + box_w + Inches(0.1))
            sub_w = int(theme.CONTENT_RIGHT - sub_left)
            if sub_w > 0:
                add_textbox(slide, sublabel, sub_left, y, sub_w, box_h,
                            font_size=Pt(10), color=theme.SUBTEXT_COLOR)

        if i < n - 1:
            arr_x = int(box_left + box_w // 2)
            arr_y1 = int(y + box_h)
            arr_y2 = int(y + box_h + gap)
            arr = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT, arr_x, arr_y1, arr_x, arr_y2)
            arr.line.color.rgb = sec_color
            arr.line.width = Pt(1.0)

    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


def build_stack_diagram_annotated(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_title_bar(slide, data["title"], section)

    sec_color = theme.SECTION_COLORS.get(section, theme.ACCENT_COLOR)
    layers = data["layers"]
    n = len(layers)
    box_h = int(Inches(0.44))
    gap = int(Inches(0.04))
    start_y = int(theme.CONTENT_TOP + Inches(0.05))

    status_fills = {
        "normal":  _darken(sec_color, 0.55),
        "warning": theme.STATUS_WARNING_FILL,
        "danger":  theme.STATUS_DANGER_FILL,
        "source":  theme.STATUS_SOURCE_FILL,
    }
    status_borders = {
        "normal":  sec_color,
        "warning": theme.STATUS_WARNING_BORDER,
        "danger":  theme.STATUS_DANGER_BORDER,
        "source":  theme.STATUS_SOURCE_BORDER,
    }

    box_left = int(theme.CONTENT_LEFT)
    box_w = int(Inches(5.2))
    ann_left = int(theme.CONTENT_LEFT + Inches(5.4))
    ann_w = int(theme.CONTENT_RIGHT - ann_left)

    for i, (label, annotation, status) in enumerate(layers):
        y = int(start_y + i * (box_h + gap))
        fill_color = status_fills.get(status, status_fills["normal"])
        border_color = status_borders.get(status, sec_color)
        border_w = Pt(2.5) if status in ("danger", "warning") else Pt(1.0)

        shape = slide.shapes.add_shape(1, box_left, y, box_w, box_h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
        shape.line.color.rgb = border_color
        shape.line.width = border_w

        tf = shape.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = label
        run.font.size = Pt(14)
        run.font.color.rgb = theme.WHITE
        run.font.bold = True
        run.font.name = theme.FONT_BODY

        ann_color = theme.SUBTEXT_COLOR
        if status == "danger":
            ann_color = theme.STATUS_DANGER_TEXT
        elif status == "warning":
            ann_color = theme.STATUS_WARNING_TEXT
        elif status == "source":
            ann_color = theme.SOURCE_ANNOTATION_COLOR

        if ann_w > 0:
            add_textbox(slide, annotation,
                        ann_left, y + int(Inches(0.05)), ann_w, box_h,
                        font_size=Pt(13), color=ann_color)

    if "note" in data:
        add_note(slide, data["note"])
    if data.get("_slide_num"):
        add_footer_bar(slide, data["_slide_num"], section)


def build_references(slide, data):
    set_slide_background(slide, theme.BG_COLOR)
    section = data.get("section", 0)
    add_title_bar(slide, data.get("qanda", "Q&A") + " / 參考資料", section)

    txBox = slide.shapes.add_textbox(
        int(theme.CONTENT_LEFT + Inches(0.15)),
        int(theme.CONTENT_TOP + Inches(0.15)),
        int(theme.CONTENT_WIDTH - Inches(0.3)),
        int(theme.CONTENT_HEIGHT - Inches(0.3))
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
        run.font.size = Pt(14)
        run.font.color.rgb = theme.ACCENT_COLOR
        run.font.name = theme.FONT_BODY
        p.space_before = Pt(5)

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
