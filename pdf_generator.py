"""
HTML mirror for Streamlit preview; PDF export uses ReportLab canvas only (no xhtml2pdf).
"""
import io
import base64
import html as html_lib
import os
import zipfile
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.pdfmetrics import stringWidth as pdf_stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas

from logic_engine import excel_text_str, qty_gt_one, resolve_row_image_bytes


def _register_pdf_fonts():
    """Register a Chinese-capable font; fallback safely when unavailable."""
    windir = os.environ.get("WINDIR", r"C:\Windows")
    font_dir = os.path.join(windir, "Fonts")
    simsun_candidates = [
        os.path.join(font_dir, "simsun.ttc"),
        os.path.join(font_dir, "simsun.ttf"),
    ]
    for font_path in simsun_candidates:
        try:
            pdfmetrics.registerFont(TTFont("SimSun", font_path))
            pdfmetrics.registerFont(TTFont("SimSun-Bold", font_path))
            return "SimSun", "SimSun-Bold"
        except Exception:
            continue
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        return "STSong-Light", "STSong-Light"
    except Exception:
        return "Helvetica", "Helvetica-Bold"


PDF_FONT_NAME, PDF_FONT_BOLD_NAME = _register_pdf_fonts()


def build_pdf_html_content(df, image_map, pdf_title):
    """Shared HTML template for mirror preview and PDF export."""
    target_df = df.iloc[:, 3:15].copy()
    img_col_name = target_df.columns[2]
    headers = ["ID1", "ID2", "IMG", "MAT", "SIZE", "QTY", "NAME", "ADDR1", "ADDR2", "CITY", "STATE", "ZIP"]
    if "材质" in target_df.columns:
        target_df["材质"] = target_df["材质"].replace("画芯", "Canvas")
    col_specs = [
        ("col-D", 20),
        ("col-E", 30),
        ("col-F", 105),
        ("col-G", 42),
        ("col-H", 30),
        ("col-I", 20),
        ("col-J", 45),
        ("col-K", 75),
        ("col-L", 65),
        ("col-M", 55),
        ("col-N", 25),
        ("col-O", 50),
    ]
    pt_to_px = 96 / 72
    target_height_px = int(125 * pt_to_px)

    prepared_images = {}
    for key, info in image_map.items():
        if not info or not info.get("bytes"):
            continue
        try:
            img_obj = Image.open(io.BytesIO(info["bytes"]))
            if img_obj.mode not in ("RGB", "L"):
                img_obj = img_obj.convert("RGB")
            ratio = target_height_px / float(img_obj.height) if img_obj.height else 1.0
            target_width_px = max(1, int(img_obj.width * ratio))
            img_obj = img_obj.resize((target_width_px, target_height_px))
            resized_buf = io.BytesIO()
            img_obj.save(resized_buf, format="JPEG", quality=85, optimize=True)
            prepared_images[key] = base64.b64encode(resized_buf.getvalue()).decode("ascii")
        except Exception:
            continue

    styling_css = """
    @page { size: a4 portrait; margin-top: 20pt; margin-bottom: 20pt; margin-left: 25pt; margin-right: 25pt; }
    body { font-family: "SimSun", "STSong", serif; font-weight: bold; color: #000000; }
    .pdf-container { font-size: 9pt; }
    .header-cell { font-size: 13px !important; color: #000000; }
    .page-title { text-align: center; font-size: 14pt; font-weight: bold; margin-bottom: 4pt; font-family: "SimSun", "STSong", serif; }
    .page-footer { text-align: center; font-size: 9pt; font-family: "SimSun", "STSong", serif; }
    table {
        table-layout: fixed !important;
        width: 540pt !important;
        border-collapse: collapse;
    }
    thead { display: table-header-group; }
    tbody { display: table-row-group; }
    tr { page-break-inside: avoid; }
    th, td {
        border: 0.75pt solid black !important;
        text-align: center;
        vertical-align: middle;
        font-size: 13px;
        padding: 2px;
        word-break: break-all;
        word-wrap: break-word;
        font-weight: bold;
        color: #000000;
    }
    th, td { font-family: "SimSun", "STSong", serif; }
    .col-D { width: 20pt !important; }
    .col-E { width: 30pt !important; }
    .col-F { width: 105pt !important; }
    .col-G {
        width: 42pt !important;
    }
    .col-H { width: 30pt !important; }
    .col-I { width: 20pt !important; }
    .col-J { width: 45pt !important; }
    .col-K { width: 75pt !important; }
    .col-L { width: 65pt !important; }
    .col-M { width: 55pt !important; }
    .col-N { width: 25pt !important; }
    .col-O { width: 50pt !important; }
    td.img-box { padding: 0; vertical-align: middle; }
    td.img-box img { max-width: 100pt; }
    """

    id1_count_map = target_df[target_df.columns[0]].map(excel_text_str).value_counts().to_dict()
    html_content = f"<html><head><meta charset='utf-8'><style>{styling_css}</style></head><body><div class='pdf-container'>"
    for i in range(0, len(target_df), 6):
        chunk = target_df.iloc[i : i + 6]
        page_num = (i // 6) + 1
        safe_title = html_lib.escape(excel_text_str(pdf_title))
        chunk_id1_vals = chunk[target_df.columns[0]].tolist()
        rowspan_start = {}
        rowspan_skip = set()
        run_start = 0
        while run_start < len(chunk_id1_vals):
            run_end = run_start
            while run_end + 1 < len(chunk_id1_vals) and chunk_id1_vals[run_end + 1] == chunk_id1_vals[run_start]:
                run_end += 1
            span = run_end - run_start + 1
            rowspan_start[run_start] = span
            for skip_idx in range(run_start + 1, run_end + 1):
                rowspan_skip.add(skip_idx)
            run_start = run_end + 1

        html_content += f"<div style='width: 540pt; font-size: 16pt; font-weight: bold; text-align: center; margin-bottom: 8px;'>{safe_title}</div>"
        html_content += "<table style='table-layout: fixed !important; width: 540pt !important; border-collapse: collapse;'><thead><tr class='header-row'>"
        for idx, h in enumerate(headers):
            col_cls, col_w = col_specs[idx]
            html_content += f"<th class='{col_cls} header-cell' style='width:{col_w}pt;'>{html_lib.escape(excel_text_str(h))}</th>"
        html_content += "</tr></thead><tbody>"

        for row_pos, (row_idx, row) in enumerate(chunk.iterrows()):
            html_content += "<tr class='data-row'>"
            for idx, col_name in enumerate(target_df.columns):
                col_cls, col_w = col_specs[idx]
                if idx == 0 and row_pos in rowspan_skip:
                    continue
                if col_name == img_col_name:
                    img_name_key = excel_text_str(df.at[row_idx, "图片名称"]).lower() if "图片名称" in df.columns else ""
                    img_b64 = prepared_images.get(img_name_key)
                    if not img_b64:
                        order_key = f"{excel_text_str(df.at[row_idx, 'purchase-date'])}_{excel_text_str(df.at[row_idx, '运单号'])}".lower()
                        img_b64 = prepared_images.get(order_key)
                    if img_b64:
                        td_content_img = f'<img src="data:image/jpeg;base64,{img_b64}">'
                    else:
                        td_content_img = "No image"
                    html_content += f"<td class='{col_cls} img-box' style='width:{col_w}pt;'>{td_content_img}</td>"
                elif col_name == "材质":
                    mat_val = excel_text_str(row[col_name])
                    is_modified = mat_val.strip() != "Canvas"
                    bg_color = "#FFFF00" if is_modified else "transparent"
                    mat_text = html_lib.escape(mat_val)

                    html_content += (
                        f'<td class="col-G" style="width: 42pt; background-color: {bg_color}; '
                        f'text-align: center; border: 0.5pt solid black; vertical-align: middle; color: #000000; font-weight: bold;">'
                        f'<div style="width: 42pt; word-wrap: break-word; word-break: break-all; '
                        f'font-size: 13px; line-height: 1.1; color: #000000; font-weight: bold;">'
                        f'{mat_text}</div></td>'
                    )
                    continue
                else:
                    td_content = html_lib.escape(excel_text_str(row[col_name]))
                    is_qty_alert = col_name == "数量" and qty_gt_one(row[col_name])
                    is_id1_merge_alert = idx == 0 and id1_count_map.get(excel_text_str(row[col_name]), 0) > 1
                    div_style = "word-wrap: break-word; width: 100%; font-size: 13px; color: #000000; font-weight: bold;"
                    td_style = f"width:{col_w}pt;"
                    td_cls = col_cls
                    if is_qty_alert:
                        div_style += " background-color: yellow; font-weight: bold;"
                        td_style += " background-color: yellow; font-weight: bold;"
                    if is_id1_merge_alert:
                        div_style += " background-color: #FFFF00; font-weight: bold;"
                        td_style += " background-color: #FFFF00; font-weight: bold;"
                    if idx == 0:
                        td_style += " vertical-align: middle; text-align: center;"
                    td_content = f"<div style='{div_style}'>{td_content}</div>"
                    if idx == 0:
                        html_content += (
                            f"<td class='{td_cls}' rowspan='{rowspan_start.get(row_pos, 1)}' style='{td_style}'>"
                            f"{td_content}</td>"
                        )
                    else:
                        html_content += f"<td class='{td_cls}' style='{td_style}'>{td_content}</td>"
            html_content += "</tr>"
        html_content += "</tbody></table>"
        html_content += f"<div class='page-footer'>Page {page_num}</div>"
        if i + 6 < len(target_df):
            html_content += "<pdf:nextpage />"
    html_content += "</div></body></html>"

    return html_content


# --- Native PDF (ReportLab canvas): pixel-accurate, no HTML/CSS ---

PAGE_W, PAGE_H = A4
MARGIN = 36
TITLE_BLOCK = 32
HEADER_H = 16
ROW_H = 120
MAX_ROWS_PAGE = 6
COL_WIDTHS = [20, 30, 105, 42, 30, 20, 45, 75, 25, 55, 25, 50]
HEADERS_EN = ["ID1", "ID2", "IMG", "MAT", "SIZE", "QTY", "NAME", "ADDR1", "ADDR2", "CITY", "STATE", "ZIP"]


def _col_x(j):
    return MARGIN + sum(COL_WIDTHS[:j])


def _wrap_text(text, max_width, font_name, font_size, c=None):
    """
    Excel-like physical wrapping: split by measured width using canvas.stringWidth
    (or pdf_stringWidth when canvas is omitted).
    """
    s = excel_text_str(text)
    if not s:
        return [""]

    def sw(part):
        if not part:
            return 0.0
        if c is not None:
            return c.stringWidth(part, font_name, font_size)
        return pdf_stringWidth(part, font_name, font_size)

    if max_width <= 0:
        return [s]

    lines = []
    i = 0
    n = len(s)
    while i < n:
        j = i + 1
        while j <= n and sw(s[i:j]) <= max_width:
            j += 1
        if j == i + 1 and sw(s[i : i + 1]) > max_width:
            lines.append(s[i])
            i += 1
        else:
            lines.append(s[i : j - 1])
            i = j - 1
    return lines if lines else [""]


def _draw_wrapped_centred_column(
    c, x_left, cell_w, y_cell_bot, cell_h, text, font_name, font_size, prefer_single_line=False
):
    """
    D–O text cells: measured wrap + geometric vertical center.
    start_y = y_cell_bot + (cell_h + total_text_h) / 2 - ascent
    """
    pad = 4.0
    max_w = max(1.0, cell_w - 2 * pad)
    c.setFillColor(colors.black)
    content = excel_text_str(text)
    lines = []
    use_font_size = font_size
    if prefer_single_line and content:
        # Keep ID2 as single-line whenever possible by shrinking font first.
        while use_font_size > 6 and c.stringWidth(content, font_name, use_font_size) > max_w:
            use_font_size -= 0.5
        if c.stringWidth(content, font_name, use_font_size) <= max_w:
            lines = [content]
        else:
            use_font_size = font_size
    c.setFont(font_name, use_font_size)
    if not lines:
        lines = _wrap_text(content, max_w, font_name, use_font_size, c)
    line_h = use_font_size * 1.18
    ascent = use_font_size * 0.72
    total_text_h = len(lines) * line_h
    start_y = y_cell_bot + (cell_h + total_text_h) / 2 - ascent
    cx = x_left + cell_w / 2
    for idx, line in enumerate(lines):
        ty = start_y - idx * line_h
        if ty < y_cell_bot + 1:
            break
        c.drawCentredString(cx, ty, line)


def _rowspan_maps(chunk_vals):
    """chunk_vals: list of ID1 cell strings in page order (already top-to-bottom)."""
    rowspan_start = {}
    rowspan_skip = set()
    rs = 0
    while rs < len(chunk_vals):
        re = rs
        while re + 1 < len(chunk_vals) and chunk_vals[re + 1] == chunk_vals[rs]:
            re += 1
        sp = re - rs + 1
        rowspan_start[rs] = sp
        for sk in range(rs + 1, re + 1):
            rowspan_skip.add(sk)
        rs = re + 1
    return rowspan_start, rowspan_skip


def _mat_display(val):
    s = excel_text_str(val).replace("画芯", "Canvas")
    return s


def _draw_native_pdf(df, image_map, pdf_title):
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    n = len(df)
    if n == 0:
        c.setFillColor(colors.black)
        c.setFont(PDF_FONT_NAME, 12)
        c.drawString(MARGIN, PAGE_H / 2, "No data")
        c.save()
        return buf.getvalue()

    display_cols = list(df.columns[3:15])
    img_col_name = display_cols[2]

    id1_values = [excel_text_str(df.iloc[i, 3]) for i in range(n)]
    id1_count_map = {}
    for val in id1_values:
        id1_count_map[val] = id1_count_map.get(val, 0) + 1
    global_page = 0
    for start in range(0, n, MAX_ROWS_PAGE):
        global_page += 1
        chunk_len = min(MAX_ROWS_PAGE, n - start)
        page_ids = [id1_values[start + i] for i in range(chunk_len)]
        page_id1_runs = []
        run_start = 0
        while run_start < chunk_len:
            run_end = run_start
            while run_end + 1 < chunk_len and page_ids[run_end + 1] == page_ids[run_start]:
                run_end += 1
            page_id1_runs.append((run_start, run_end, page_ids[run_start]))
            run_start = run_end + 1

        y_top = PAGE_H - MARGIN
        c.setFillColor(colors.black)
        c.setFont(PDF_FONT_BOLD_NAME, 14)
        c.drawCentredString(PAGE_W / 2, y_top - 18, excel_text_str(pdf_title))

        y_header_top = y_top - TITLE_BLOCK
        y_header_bottom = y_header_top - HEADER_H
        for j in range(12):
            x0 = _col_x(j)
            c.setStrokeColor(colors.black)
            c.setFillColor(colors.white)
            c.rect(x0, y_header_bottom, COL_WIDTHS[j], HEADER_H, stroke=1, fill=1)
            c.setFillColor(colors.black)
            c.setFont(PDF_FONT_BOLD_NAME, 10.5)
            c.drawCentredString(x0 + COL_WIDTHS[j] / 2, y_header_bottom + HEADER_H / 2 - 3, HEADERS_EN[j])

        for row_pos in range(chunk_len):
            pos = start + row_pos
            full_row = df.iloc[pos]
            y_row_top = y_header_bottom - row_pos * ROW_H
            y_row_bot = y_row_top - ROW_H

            for ci, col_name in enumerate(display_cols):
                x0 = _col_x(ci)
                w = COL_WIDTHS[ci]

                c.setStrokeColor(colors.black)
                c.setLineWidth(0.75)
                cell_h = ROW_H
                y_cell_bot = y_row_bot

                if ci == 0:
                    # ID1 column is rendered as merged blocks after row rendering.
                    continue

                if col_name == "材质":
                    mat_raw = full_row[col_name]
                    mat_show = _mat_display(mat_raw)
                    c.setFillColor(colors.yellow if mat_show.strip() != "Canvas" else colors.white)
                    c.rect(x0, y_cell_bot, w, cell_h, stroke=1, fill=1)
                    _draw_wrapped_centred_column(
                        c, x0, w, y_cell_bot, cell_h, mat_show, PDF_FONT_BOLD_NAME, 10
                    )
                    continue

                is_qty = col_name == "数量" and qty_gt_one(full_row[col_name])
                c.setFillColor(colors.yellow if is_qty else colors.white)
                c.rect(x0, y_cell_bot, w, cell_h, stroke=1, fill=1)

                if col_name == img_col_name:
                    img_bytes = resolve_row_image_bytes(full_row, image_map)
                    if img_bytes:
                        try:
                            pil_im = Image.open(io.BytesIO(img_bytes))
                            if pil_im.mode not in ("RGB", "L"):
                                pil_im = pil_im.convert("RGB")
                            iw, ih = pil_im.size
                            mx_w, mx_h = w - 4, cell_h - 4
                            scale = min(mx_w / iw, mx_h / ih, 1.0)
                            dw, dh = iw * scale, ih * scale
                            xi = x0 + (w - dw) / 2
                            yi = y_cell_bot + (cell_h - dh) / 2
                            bbuf = io.BytesIO()
                            pil_im.save(bbuf, format="JPEG", quality=85)
                            bbuf.seek(0)
                            c.drawImage(ImageReader(bbuf), xi, yi, width=dw, height=dh, mask="auto")
                        except Exception:
                            _draw_wrapped_centred_column(
                                c, x0, w, y_cell_bot, cell_h, "No image", PDF_FONT_BOLD_NAME, 10
                            )
                    else:
                        _draw_wrapped_centred_column(
                            c, x0, w, y_cell_bot, cell_h, "No image", PDF_FONT_BOLD_NAME, 10
                        )
                else:
                    txt = excel_text_str(full_row[col_name])
                    _draw_wrapped_centred_column(
                        c, x0, w, y_cell_bot, cell_h, txt, PDF_FONT_BOLD_NAME, 10, prefer_single_line=(ci == 1)
                    )

        # Draw ID1 merged blocks for this page.
        id1_x = _col_x(0)
        id1_w = COL_WIDTHS[0]
        c.setStrokeColor(colors.black)
        c.setLineWidth(0.75)
        for run_start, run_end, run_id in page_id1_runs:
            y_run_top = y_header_bottom - run_start * ROW_H
            y_run_bot = y_header_bottom - (run_end + 1) * ROW_H
            is_merged_id1 = id1_count_map.get(run_id, 0) > 1
            c.setFillColor(colors.yellow if is_merged_id1 else colors.white)
            c.rect(id1_x, y_run_bot, id1_w, y_run_top - y_run_bot, stroke=1, fill=1)
            c.setFillColor(colors.black)
            c.setFont(PDF_FONT_BOLD_NAME, 10.5)
            c.drawCentredString(id1_x + id1_w / 2.0, y_run_bot + (y_run_top - y_run_bot) / 2.0 - 3, run_id)

        c.setFillColor(colors.black)
        c.setFont(PDF_FONT_BOLD_NAME, 10)
        c.drawString(MARGIN, MARGIN / 2, f"Page {global_page}")
        if start + MAX_ROWS_PAGE < n:
            c.showPage()

    c.save()
    return buf.getvalue()


def generate_pdf_with_images(df, image_map, pdf_title):
    return _draw_native_pdf(df, image_map, pdf_title)


PDF_IMAGE_DPI = 150
PDF_IMAGE_SIZE = (1241, 1754)


def pdf_to_images_zip(pdf_bytes, pdf_title):
    """Convert final PDF pages to 150 DPI, 1241x1754 PNG files in an in-memory ZIP archive."""
    import fitz

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    zip_buffer = io.BytesIO()
    render_matrix = fitz.Matrix(PDF_IMAGE_DPI / 72, PDF_IMAGE_DPI / 72)

    try:
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                pix = page.get_pixmap(matrix=render_matrix)
                mode = "RGBA" if pix.alpha else "RGB"
                img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
                if mode == "RGBA":
                    img = img.convert("RGB")
                img = img.resize(PDF_IMAGE_SIZE, Image.Resampling.LANCZOS)

                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format="PNG", dpi=(PDF_IMAGE_DPI, PDF_IMAGE_DPI))
                img_data = img_byte_arr.getvalue()

                seq = str(page_num + 1).zfill(2)
                img_name = f"0-{seq}-{pdf_title}.png"
                zip_file.writestr(img_name, img_data)
    finally:
        doc.close()

    return zip_buffer.getvalue()
