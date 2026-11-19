import io
import re
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.dml.color import RGBColor

MIN_LINE_THICKNESS = Pt(0.75)
IMAGE_SCALE = 3
TEXT_VISIBILITY_SCALE = 4
TEXT_VISIBILITY_THRESHOLD = 0.01


def _is_text_visible(page, rect, color_int):
    """Some PDFs (template-generated dashboards) draw placeholder text -
    e.g. a "no data" fallback - at a fixed slot and then clip or cover it
    wherever real content exists, without changing its reported opacity.
    PyMuPDF's text/drawing extraction APIs don't expose that clip state,
    so the only reliable signal is rendering the live page: if the span's
    own ink color barely shows up in its own bbox, it isn't really there."""
    if rect.width <= 0 or rect.height <= 0:
        return True
    pix = page.get_pixmap(clip=rect, matrix=fitz.Matrix(TEXT_VISIBILITY_SCALE, TEXT_VISIBILITY_SCALE))
    target = ((color_int >> 16) & 255, (color_int >> 8) & 255, color_int & 255)
    samples = pix.samples
    total = pix.width * pix.height
    if total == 0:
        return True
    close = 0
    for i in range(0, len(samples), 3):
        if abs(samples[i] - target[0]) + abs(samples[i + 1] - target[1]) + abs(samples[i + 2] - target[2]) < 60:
            close += 1
    return (close / total) >= TEXT_VISIBILITY_THRESHOLD


def _rgb(color):
    """Accepts a (r,g,b) float tuple (0-1) or a 24-bit int (0xRRGGBB)."""
    if isinstance(color, (tuple, list)):
        return RGBColor(int(color[0] * 255), int(color[1] * 255), int(color[2] * 255))
    r = (color >> 16) & 255
    g = (color >> 8) & 255
    b = color & 255
    return RGBColor(r, g, b)


def _dash_style(dashes):
    """PyMuPDF reports dashes as a PDF dash-array string, e.g. "[ 1 ] 0".
    An empty/missing array means solid; short segments read as a dotted
    line, longer ones as dashed."""
    if not dashes:
        return None
    nums = [float(n) for n in re.findall(r"[\d.]+", dashes.split("]")[0])]
    if not nums or max(nums) == 0:
        return None
    return MSO_LINE_DASH_STYLE.ROUND_DOT if max(nums) <= 2 else MSO_LINE_DASH_STYLE.DASH


def _flatten_cubic(p0, p1, p2, p3, n=8):
    pts = []
    for i in range(1, n + 1):
        t = i / n
        mt = 1 - t
        x = mt**3 * p0.x + 3 * mt**2 * t * p1.x + 3 * mt * t**2 * p2.x + t**3 * p3.x
        y = mt**3 * p0.y + 3 * mt**2 * t * p1.y + 3 * mt * t**2 * p2.y + t**3 * p3.y
        pts.append((x, y))
    return pts


def _path_vertices(items):
    """Walk a drawing's path commands into an ordered polygon, flattening
    any bezier curves into short line segments."""
    pts = []
    for it in items:
        kind = it[0]
        if kind == "l":
            _, p0, p1 = it
            if not pts:
                pts.append((p0.x, p0.y))
            pts.append((p1.x, p1.y))
        elif kind == "c":
            _, p0, p1, p2, p3 = it
            if not pts:
                pts.append((p0.x, p0.y))
            pts.extend(_flatten_cubic(p0, p1, p2, p3))
        elif kind == "re":
            r = it[1]
            pts.extend([(r.x0, r.y0), (r.x1, r.y0), (r.x1, r.y1), (r.x0, r.y1)])
        elif kind == "qu":
            q = it[1]
            pts.extend([(p.x, p.y) for p in (q.ul, q.ur, q.lr, q.ll)])
    deduped = []
    for pt in pts:
        if not deduped or abs(pt[0] - deduped[-1][0]) > 0.1 or abs(pt[1] - deduped[-1][1]) > 0.1:
            deduped.append(pt)
    if len(deduped) > 1 and abs(deduped[0][0] - deduped[-1][0]) < 0.1 and abs(deduped[0][1] - deduped[-1][1]) < 0.1:
        deduped.pop()
    return deduped


def _is_axis_rect(verts, tol=0.5):
    if len(verts) != 4:
        return False
    for i in range(4):
        x0, y0 = verts[i]
        x1, y1 = verts[(i + 1) % 4]
        if abs(x0 - x1) > tol and abs(y0 - y1) > tol:
            return False
    return True


def _rounded_rect_adjustment(items, rect):
    """Heuristic for a rectangle whose corners are short bezier arcs: use
    the arcs' chord length as the corner radius, expressed as the ratio
    MSO_SHAPE.ROUNDED_RECTANGLE expects (fraction of the shorter side)."""
    curves = [it for it in items if it[0] == "c"]
    lines = [it for it in items if it[0] == "l"]
    if not curves or not lines or len(curves) > 4 or len(lines) > 4:
        return None
    min_side = min(rect.width, rect.height)
    if min_side <= 0:
        return None
    chords = [((it[1].x - it[4].x) ** 2 + (it[1].y - it[4].y) ** 2) ** 0.5 for it in curves]
    radius = (sum(chords) / len(chords)) / 2
    return max(0.0, min(radius / min_side, 0.5))


FREEFORM_SUBUNITS_PER_POINT = 100  # local unit = 1/100 pt


def _add_freeform(slide, verts):
    if len(verts) < 3:
        return None
    # build_freeform rounds every vertex to the nearest whole LOCAL unit
    # before scaling to EMU. Passing raw point values as local units loses
    # any sub-point geometry (two coordinates less than 1pt apart can round
    # to the same point and collapse the shape to zero width/height - e.g.
    # thin 0.75pt border-edge rectangles). Scaling up to hundredths of a
    # point first keeps that precision through the rounding step.
    local = [(x * FREEFORM_SUBUNITS_PER_POINT, y * FREEFORM_SUBUNITS_PER_POINT) for x, y in verts]
    fb = slide.shapes.build_freeform(
        start_x=local[0][0], start_y=local[0][1], scale=Pt(1) / FREEFORM_SUBUNITS_PER_POINT)
    fb.add_line_segments(local[1:], close=True)
    return fb.convert_to_shape()


def _add_drawing_shape(slide, items, rect):
    """Classify a PDF vector drawing's path and render the closest-matching
    PowerPoint shape: rectangle, oval, rounded rectangle, or an arbitrary
    freeform polygon (triangles, stars, rotated/irregular shapes, ...)."""
    items = items or []
    kinds = [it[0] for it in items]
    x, y = Pt(rect.x0), Pt(rect.y0)
    w, h = Pt(max(rect.width, 0.01)), Pt(max(rect.height, 0.01))

    if not kinds or kinds == ["re"]:
        return slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)

    if kinds and all(k == "c" for k in kinds) and len(kinds) >= 3:
        return slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)

    if set(kinds) == {"l"}:
        verts = _path_vertices(items)
        if _is_axis_rect(verts):
            return slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
        return _add_freeform(slide, verts) or slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)

    if set(kinds) <= {"l", "c"}:
        radius = _rounded_rect_adjustment(items, rect)
        if radius is not None:
            shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
            shape.adjustments[0] = radius
            return shape
        verts = _path_vertices(items)
        return _add_freeform(slide, verts) or slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)

    verts = _path_vertices(items)
    return _add_freeform(slide, verts) or slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)


LINK_MATCH_RATIO = 0.75
LINK_OVERLAY_MAX_SIDE = 120


def _rect_area(r):
    return max(r.width, 0) * max(r.height, 0)


def _overlap_ratio(a, b):
    ix0, iy0 = max(a.x0, b.x0), max(a.y0, b.y0)
    ix1, iy1 = min(a.x1, b.x1), min(a.y1, b.y1)
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    inter = (ix1 - ix0) * (iy1 - iy0)
    smaller = min(_rect_area(a), _rect_area(b))
    return inter / smaller if smaller > 0 else 0.0


def _best_link_target(link_rect, candidates, max_area):
    """Among drawing/text elements already on the slide, pick the smallest
    one that substantially contains (or is contained by) the link's rect -
    i.e. the most specific panel/line this link actually annotates. Huge
    elements (e.g. the full-page background) are excluded so a link never
    falls back to "clickable everywhere"."""
    best = None
    best_area = None
    for rect, payload in candidates:
        area = _rect_area(rect)
        if area <= 0 or area > max_area:
            continue
        if _overlap_ratio(link_rect, rect) < LINK_MATCH_RATIO:
            continue
        if best is None or area < best_area:
            best, best_area = payload, area
    return best


def _font_name_and_style(span_font, flags):
    # PyMuPDF font names look like "Calibri,Bold" or "Arial-BoldMT" - strip
    # the style suffix so PowerPoint resolves the real family, and derive
    # bold/italic from the name plus the span's flag bits (bit1=italic,
    # bit4=bold) since the name alone isn't always reliable.
    name = span_font.split(",")[0].split("-")[0].strip()
    lower = span_font.lower()
    bold = bool(flags & 16) or "bold" in lower
    italic = bool(flags & 2) or "italic" in lower or "oblique" in lower
    return name, bold, italic


def pdf_to_pptx(pdf_path, out_path):
    doc = fitz.open(pdf_path)
    prs = Presentation()
    first = doc[0].rect
    prs.slide_width = Pt(first.width)
    prs.slide_height = Pt(first.height)
    blank = prs.slide_layouts[6]

    # Slides are created up front so internal "jump to page" links can
    # target any slide, including ones not yet populated.
    slides = [prs.slides.add_slide(blank) for _ in doc]

    for page_index, page in enumerate(doc):
        slide = slides[page_index]
        page_area = page.rect.width * page.rect.height
        link_candidates = []  # (rect, shape_or_runs) - panels and text lines a link can attach to

        # 1. Vector drawings (fills + strokes), drawn first so they sit
        # behind images and text, in their original z-order.
        for d in page.get_drawings():
            rect = d.get("rect")
            if not rect or rect.width <= 0 and rect.height <= 0:
                continue
            dtype = d.get("type", "")

            if "f" in dtype and d.get("fill") and d.get("fill_opacity", 1.0) > 0:
                shape = _add_drawing_shape(slide, d.get("items"), rect)
                shape.fill.solid()
                shape.fill.fore_color.rgb = _rgb(d["fill"])
                shape.line.fill.background()
                shape.shadow.inherit = False
                link_candidates.append((rect, ("shape", shape)))

            if "s" in dtype and d.get("color") and d.get("stroke_opacity", 1.0) > 0:
                width_pt = d.get("width") or 1.0
                dash = _dash_style(d.get("dashes"))
                items = d.get("items") or []
                is_horizontal = rect.height <= rect.width * 0.05
                is_vertical = rect.width <= rect.height * 0.05
                if len(items) == 1 and items[0][0] == "l":
                    # A lone line segment - use its exact endpoints so
                    # diagonal lines aren't flattened into a bounding box.
                    _, p0, p1 = items[0]
                    connector = slide.shapes.add_connector(
                        MSO_CONNECTOR.STRAIGHT, Pt(p0.x), Pt(p0.y), Pt(p1.x), Pt(p1.y))
                    connector.line.color.rgb = _rgb(d["color"])
                    connector.line.width = Pt(width_pt)
                    if dash:
                        connector.line.dash_style = dash
                    connector.shadow.inherit = False
                elif is_horizontal or is_vertical:
                    # Real connector lines (not filled bars) so dash/dot
                    # patterns can be applied.
                    cy = (rect.y0 + rect.y1) / 2
                    cx = (rect.x0 + rect.x1) / 2
                    if is_horizontal:
                        connector = slide.shapes.add_connector(
                            MSO_CONNECTOR.STRAIGHT, Pt(rect.x0), Pt(cy), Pt(rect.x1), Pt(cy))
                    else:
                        connector = slide.shapes.add_connector(
                            MSO_CONNECTOR.STRAIGHT, Pt(cx), Pt(rect.y0), Pt(cx), Pt(rect.y1))
                    connector.line.color.rgb = _rgb(d["color"])
                    connector.line.width = Pt(width_pt)
                    if dash:
                        connector.line.dash_style = dash
                    connector.shadow.inherit = False
                else:
                    shape = _add_drawing_shape(slide, items, rect)
                    shape.fill.background()
                    shape.line.color.rgb = _rgb(d["color"])
                    shape.line.width = Pt(width_pt)
                    if dash:
                        shape.line.dash_style = dash
                    shape.shadow.inherit = False
                    link_candidates.append((rect, ("shape", shape)))

        # 2. Images - rasterized directly from the live page rather than
        # re-extracting the embedded XObject bytes. get_image_rects() only
        # reports an image's nominal placement rect, not any clip path
        # active when it was painted, so a raw extraction can resurrect
        # content that is actually clipped away (e.g. a hidden "no data"
        # placeholder icon sitting under real content elsewhere on the
        # page). Re-rendering the page at that exact rect bakes in the
        # correct clipping, z-order and compositing for free.
        for img in page.get_images(full=True):
            xref = img[0]
            for r in page.get_image_rects(xref):
                if r.width <= 0 or r.height <= 0:
                    continue
                pix = page.get_pixmap(clip=r, matrix=fitz.Matrix(IMAGE_SCALE, IMAGE_SCALE), alpha=True)
                slide.shapes.add_picture(
                    io.BytesIO(pix.tobytes("png")),
                    Pt(r.x0), Pt(r.y0), Pt(r.width), Pt(r.height))

        # 3. Text lines as editable text boxes. A PDF "block" can contain
        # lines at different x-offsets (e.g. a number column next to a
        # wrapped label column), so each line gets its own box positioned
        # at the line's own bbox rather than stacking lines inside one
        # block-wide box, which would destroy that side-by-side layout.
        for block in page.get_text("dict")["blocks"]:
            if block["type"] != 0:  # skip image blocks
                continue
            for line in block["lines"]:
                visible_spans = [
                    span for span in line["spans"]
                    if span["text"].strip() == "" or _is_text_visible(page, fitz.Rect(span["bbox"]), span["color"])
                ]
                if not visible_spans:
                    continue
                x0, y0, x1, y1 = line["bbox"]
                tb = slide.shapes.add_textbox(
                    Pt(x0), Pt(y0), Pt(x1 - x0 + 2), Pt(y1 - y0))
                tf = tb.text_frame
                tf.word_wrap = False
                tf.margin_left = tf.margin_right = 0
                tf.margin_top = tf.margin_bottom = 0
                p = tf.paragraphs[0]
                runs = []
                for span in visible_spans:
                    run = p.add_run()
                    run.text = span["text"]
                    run.font.size = Pt(span["size"])
                    name, bold, italic = _font_name_and_style(span["font"], span["flags"])
                    run.font.name = name
                    run.font.bold = bold
                    run.font.italic = italic
                    run.font.color.rgb = _rgb(span["color"])
                    runs.append(run)
                link_candidates.append((fitz.Rect(x0, y0, x1, y1), ("runs", runs)))

        # 4. Hyperlinks. Rather than stacking a new invisible clickable
        # rectangle over every PDF link annotation - which for
        # template-exported dashboards can mean dozens of overlapping
        # whole-panel/whole-sidebar link regions burying the real text
        # underneath - attach the link directly to whichever text line or
        # background panel most specifically matches its rect. Only create
        # a standalone overlay as a last resort, and only when it's small
        # enough not to block anything (e.g. a bare icon with no panel or
        # text under it).
        for link in page.get_links():
            rect = link.get("from")
            if not rect or rect.width <= 0 or rect.height <= 0:
                continue
            kind = link.get("kind")
            address = link.get("uri") if kind == fitz.LINK_URI else None
            target_slide = (
                slides[link["page"]]
                if kind == fitz.LINK_GOTO and link.get("page", -1) in range(len(slides))
                else None
            )
            if not address and target_slide is None:
                continue

            target = _best_link_target(rect, link_candidates, max_area=0.5 * page_area)
            if target and target[0] == "runs" and address:
                for run in target[1]:
                    run.hyperlink.address = address
                continue
            if target and target[0] == "shape":
                if address:
                    target[1].click_action.hyperlink.address = address
                elif target_slide is not None:
                    target[1].click_action.target_slide = target_slide
                continue

            if rect.width <= LINK_OVERLAY_MAX_SIDE and rect.height <= LINK_OVERLAY_MAX_SIDE:
                overlay = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, Pt(rect.x0), Pt(rect.y0),
                    Pt(rect.width), Pt(rect.height))
                overlay.fill.background()
                overlay.line.fill.background()
                overlay.shadow.inherit = False
                if address:
                    overlay.click_action.hyperlink.address = address
                elif target_slide is not None:
                    overlay.click_action.target_slide = target_slide

    prs.save(out_path)


pdf_to_pptx(
    r"\expert_summary_form_stoke.pdf",
    r"\expert_summary_form_stoke.pptx")
