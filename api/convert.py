"""
Conversiones de PDF a PowerPoint y a Excel con PyMuPDF.

PDF a PowerPoint: cada diapositiva lleva de fondo la página sin el texto (imágenes, colores y
formas tal cual) y encima el texto en cuadros editables, en su sitio y con su fuente, tamaño,
negrita, cursiva y color, leídos directamente del PDF. Con mode='image' cada diapositiva es
solo la imagen de la página, con el texto en las notas del orador.

PDF a Excel: una hoja por página. Las tablas que detecta PyMuPDF van como tablas; el resto del
texto, una fila por línea, con una celda por cada bloque separado por un hueco grande.
"""

import io
import math
import re
from dataclasses import dataclass, field

import pymupdf
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.util import Emu, Pt


class UserError(Exception):
    """Error que se puede explicar a la persona (se devuelve con su mensaje)."""


def open_pdf(data: bytes, max_pages: int) -> pymupdf.Document:
    doc = pymupdf.open(stream=data, filetype='pdf')
    if doc.needs_pass:
        raise UserError('El PDF está protegido con contraseña. Desbloquéalo antes de convertirlo.')
    if doc.page_count > max_pages:
        raise UserError(f'El PDF tiene {doc.page_count} páginas; el máximo es {max_pages}.')
    if doc.page_count == 0:
        raise UserError('El PDF no tiene páginas.')
    return doc


# ---------- Texto de la página ----------

# Caracteres que no admite el XML de Office.
_BAD_XML = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]')

_FONTS = {
    'Helvetica': 'Arial', 'Arial': 'Arial', 'ArialNarrow': 'Arial Narrow', 'Times': 'Times New Roman',
    'TimesNewRoman': 'Times New Roman', 'TimesRoman': 'Times New Roman', 'Courier': 'Courier New',
    'CourierNew': 'Courier New', 'Symbol': 'Symbol', 'Calibri': 'Calibri', 'CalibriLight': 'Calibri Light',
    'Cambria': 'Cambria', 'Candara': 'Candara', 'Consolas': 'Consolas', 'Georgia': 'Georgia',
    'Verdana': 'Verdana', 'Tahoma': 'Tahoma', 'TrebuchetMS': 'Trebuchet MS', 'Garamond': 'Garamond',
    'BookAntiqua': 'Book Antiqua', 'CenturyGothic': 'Century Gothic', 'SegoeUI': 'Segoe UI', 'Aptos': 'Aptos',
    'Roboto': 'Roboto', 'OpenSans': 'Open Sans', 'Montserrat': 'Montserrat', 'Lato': 'Lato', 'Poppins': 'Poppins',
    'SourceSansPro': 'Source Sans Pro',
    # Fuentes libres con las mismas medidas que las de Office (LibreOffice, Google Docs).
    'LiberationSans': 'Arial', 'Arimo': 'Arial', 'LiberationSerif': 'Times New Roman', 'Tinos': 'Times New Roman',
    'LiberationMono': 'Courier New', 'Cousine': 'Courier New', 'Carlito': 'Calibri', 'Caladea': 'Cambria',
}


def font_face(raw: str, flags: int) -> str:
    """Nombre de la fuente del PDF → nombre que entiende PowerPoint."""
    name = re.split(r'[-,]', re.sub(r'^[A-Z]{6}\+', '', raw))[0]
    name = re.sub(r'(PSMT|MT|PS|Std|Pro|LT)$', '', name)
    if name in _FONTS:
        return _FONTS[name]
    if re.fullmatch(r'[A-Za-z][A-Za-z0-9]+', name) and not re.match(r'(F\d+|T\d|TT\d|g_)', name):
        return re.sub(r'([a-z])([A-Z])', r'\1 \2', name)
    # Sin nombre reconocible: según el tipo de letra (bit 4 = con remates, 8 = monoespaciada).
    return 'Courier New' if flags & 8 else 'Times New Roman' if flags & 4 else 'Arial'


@dataclass
class Frag:
    text: str
    x: float
    y: float  # línea base
    w: float
    size: float
    angle: int
    bold: bool
    italic: bool
    font: str
    color: str


def page_frags(page: pymupdf.Page, hidden: bool = False) -> list[Frag]:
    """
    Fragmentos de texto en coordenadas de la página tal como se ve (con su giro). Con hidden,
    también el texto invisible (la capa de un OCR), que en PowerPoint quedaría duplicado sobre la imagen.
    """
    m = page.rotation_matrix
    zero = pymupdf.Point(0, 0) * m
    frags = []
    for block in page.get_text('dict', flags=pymupdf.TEXT_PRESERVE_WHITESPACE)['blocks']:
        for line in block.get('lines', []):
            d = pymupdf.Point(line['dir']) * m - zero
            angle = round(math.degrees(math.atan2(d.y, d.x)))
            for s in line['spans']:
                text = _BAD_XML.sub('', s['text'])
                if not text.strip() or s['size'] < 1 or (s.get('alpha', 255) == 0 and not hidden):
                    continue
                o = pymupdf.Point(s['origin']) * m
                r = pymupdf.Rect(s['bbox']) * m
                flags = s['flags']
                frags.append(Frag(
                    text=text, x=o.x, y=o.y, w=max(r.width, r.height) if abs(angle) > 1 else r.width,
                    size=s['size'], angle=0 if abs(angle) <= 1 else angle,
                    bold=bool(flags & 16) or bool(re.search(r'bold|black|heavy|semibold|demi', s['font'], re.I)),
                    italic=bool(flags & 2) or bool(re.search(r'italic|oblique', s['font'], re.I)),
                    font=font_face(s['font'], flags), color=f"{s['color']:06X}",
                ))
    return frags


@dataclass
class Seg:
    frags: list[Frag]
    x: float
    right: float
    y: float
    size: float


@dataclass
class Box:
    segs: list[Seg] = field(default_factory=list)
    x: float = 0
    right: float = 0
    size: float = 0


def to_segments(frags: list[Frag]) -> list[Seg]:
    """Fragmentos de la misma línea, cortados donde hay un hueco grande (columnas, tablas)."""
    lines: list[list[Frag]] = []
    for f in sorted((f for f in frags if not f.angle), key=lambda f: (f.y, f.x)):
        if lines and abs(f.y - lines[-1][0].y) < max(lines[-1][0].size, f.size) * 0.45:
            lines[-1].append(f)
        else:
            lines.append([f])
    segs = []
    for line in lines:
        line.sort(key=lambda f: f.x)
        cur: list[Frag] = []

        def flush():
            if cur and any(f.text.strip() for f in cur):
                segs.append(Seg(list(cur), cur[0].x, cur[-1].x + cur[-1].w, cur[0].y, max(f.size for f in cur)))
            cur.clear()

        for f in line:
            if cur:
                prev = cur[-1]
                gap = f.x - (prev.x + prev.w)
                if gap > max(prev.size, f.size) * 1.5:
                    flush()
                elif gap > f.size * 0.18 and not prev.text.endswith(' ') and not f.text.startswith(' '):
                    f.text = ' ' + f.text
            cur.append(f)
        flush()
    return segs


def to_boxes(segs: list[Seg]) -> list[Box]:
    """Une los segmentos que forman un mismo párrafo (misma columna, mismo tamaño, seguidos)."""
    boxes: list[Box] = []
    for s in sorted(segs, key=lambda s: (s.y, s.x)):
        target = None
        for b in boxes:
            last = b.segs[-1]
            gap = s.y - last.y
            same_lead = len(b.segs) < 2 or abs(gap - (last.y - b.segs[-2].y)) < s.size * 0.15
            if (abs(s.size - b.size) < 0.6 and s.size * 0.8 < gap < s.size * 1.6 and abs(s.x - b.x) < s.size * 1.5
                    and s.right > b.x and s.x < b.right and same_lead):
                target = b
                break
        if target:
            target.segs.append(s)
            target.x = min(target.x, s.x)
            target.right = max(target.right, s.right)
        else:
            boxes.append(Box([s], s.x, s.right, s.size))
    return boxes


# ---------- PDF a PowerPoint ----------

def _emu(pt: float) -> Emu:
    return Emu(int(round(pt * 12700)))


def _background(doc: pymupdf.Document, index: int, dpi: int = 150) -> pymupdf.Pixmap:
    """La página sin texto: se quita todo el texto y se dejan intactas imágenes y formas."""
    tmp = pymupdf.open()
    tmp.insert_pdf(doc, from_page=index, to_page=index)
    page = tmp[0]
    page.add_redact_annot(page.mediabox, fill=False)
    page.apply_redactions(
        images=pymupdf.PDF_REDACT_IMAGE_NONE,
        graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
        text=pymupdf.PDF_REDACT_TEXT_REMOVE,
    )
    pix = page.get_pixmap(dpi=dpi)
    tmp.close()
    return pix


def _style(run, f: Frag, k: float):
    run.text = f.text
    font = run.font
    font.size = Pt(round(f.size * k * 10) / 10)
    font.bold = f.bold
    font.italic = f.italic
    font.name = f.font
    font.color.rgb = RGBColor.from_string(f.color)


def _textbox(slide, x, y, w, h, wrap=False):
    tb = slide.shapes.add_textbox(_emu(x), _emu(y), _emu(w), _emu(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tb, tf


def pdf_to_pptx(data: bytes, mode: str, max_pages: int) -> bytes:
    doc = open_pdf(data, max_pages)
    prs = Presentation()
    # Diapositivas del tamaño de la primera página: 1 punto del PDF = 1 punto en PowerPoint.
    W, H = doc[0].rect.width, doc[0].rect.height
    prs.slide_width, prs.slide_height = _emu(W), _emu(H)
    layout = prs.slide_layouts[6]  # en blanco

    for page in doc:
        slide = prs.slides.add_slide(layout)
        r = page.rect
        k = min(W / r.width, H / r.height)  # si una página tiene otra proporción, se centra
        ox, oy = (W - r.width * k) / 2, (H - r.height * k) / 2
        frags = page_frags(page) if mode == 'editable' else []
        pix = _background(doc, page.number) if frags else page.get_pixmap(dpi=150)
        slide.shapes.add_picture(io.BytesIO(pix.tobytes('jpeg', jpg_quality=88)), _emu(ox), _emu(oy), _emu(r.width * k), _emu(r.height * k))

        if mode != 'editable':
            text = _BAD_XML.sub('', page.get_text()).strip()
            if text:
                slide.notes_slide.notes_text_frame.text = text
            continue

        for box in to_boxes(to_segments(frags)):
            n = len(box.segs)
            lead = (box.segs[-1].y - box.segs[0].y) / (n - 1) if n > 1 else box.size * 1.2
            top = box.segs[0].y - (lead * 0.86 if n > 1 else box.size)
            w = max((box.right - box.x) * 1.04 + box.size * 0.5, box.size)
            h = lead * n if n > 1 else box.size * 1.25
            _, tf = _textbox(slide, ox + box.x * k, oy + top * k, w * k, h * k)
            for i, seg in enumerate(box.segs):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                if n > 1:
                    p.line_spacing = Pt(round(lead * k * 10) / 10)
                p.space_before = p.space_after = Pt(0)
                for j, f in enumerate(seg.frags):
                    if i > 0 and j == 0:
                        f.text = f.text.lstrip()
                    _style(p.add_run(), f, k)

        # Texto girado: un cuadro por fragmento, girado alrededor de su centro.
        for f in (f for f in frags if f.angle):
            rad = math.radians(f.angle)
            cx = f.x + math.cos(rad) * (f.w / 2) + math.sin(rad) * (f.size * 0.35)
            cy = f.y + math.sin(rad) * (f.w / 2) - math.cos(rad) * (f.size * 0.35)
            w, h = f.w * 1.05 + f.size * 0.3, f.size * 1.25
            tb, tf = _textbox(slide, ox + (cx - w / 2) * k, oy + (cy - h / 2) * k, w * k, h * k)
            tb.rotation = f.angle % 360
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            _style(tf.paragraphs[0].add_run(), f, k)

    doc.close()
    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()


# ---------- PDF a Excel ----------

def to_number(s: str, lang: str):
    """Convierte «1.234,56» (es) o «1,234.56» (en) en número; si no es un número claro, deja el texto."""
    t = s.strip().replace(' ', ' ')
    if not t or len(t) > 20:
        return s
    neg = t.startswith('-')
    body = t[1:] if neg else t
    if re.fullmatch(r'0\d+', body):  # códigos con ceros delante (007, 01234): se quedan como texto
        return s
    thousands, decimal = ('.', ',') if lang == 'es' else (',', '.')
    if re.fullmatch(rf'\d{{1,3}}(\{thousands}\d{{3}})+(\{decimal}\d+)?', body) or re.fullmatch(rf'\d+(\{decimal}\d+)?', body):
        num = body.replace(thousands, '').replace(decimal, '.')
        value = float(num) if '.' in num else int(num)
        return -value if neg else value
    return s


def pdf_to_xlsx(data: bytes, lang: str, max_pages: int) -> bytes:
    doc = open_pdf(data, max_pages)
    wb = Workbook()
    wb.remove(wb.active)
    label = 'Página' if lang == 'es' else 'Page'

    for page in doc:
        ws = wb.create_sheet(f'{label} {page.number + 1}')
        try:
            tables = page.find_tables().tables
        except Exception:  # noqa: BLE001 - si falla la detección, la página va como texto
            tables = []
        boxes = [pymupdf.Rect(t.bbox) * page.rotation_matrix for t in tables]

        # Lo que no está dentro de una tabla, línea a línea (cada bloque separado, una celda).
        items = []
        for seg in to_segments(page_frags(page, hidden=True)):
            cx, cy = (seg.x + seg.right) / 2, seg.y
            if any(b.contains(pymupdf.Point(cx, cy)) for b in boxes):
                continue
            text = ''.join(f.text for f in seg.frags).strip()
            if items and items[-1][0] == 'row' and abs(items[-1][1] - seg.y) < seg.size * 0.45:
                items[-1][2].append((seg.x, text))
            else:
                items.append(['row', seg.y, [(seg.x, text)]])
        for t, b in zip(tables, boxes):
            items.append(['table', b.y0, t.extract()])
        items.sort(key=lambda it: it[1])

        row = 1
        widths: dict[int, int] = {}
        for kind, _, content in items:
            rows = [[c for _, c in sorted(content)]] if kind == 'row' else content
            for values in rows:
                for col, value in enumerate(values, start=1):
                    if value is None or value == '':
                        continue
                    text = _BAD_XML.sub('', str(value)).replace('\n', ' ').strip()
                    ws.cell(row=row, column=col, value=to_number(text, lang))
                    widths[col] = max(widths.get(col, 0), min(len(text), 60))
                row += 1
            if kind == 'table':
                row += 1  # una fila en blanco después de cada tabla
        for col, wdt in widths.items():
            ws.column_dimensions[get_column_letter(col)].width = max(8, wdt + 2)

    doc.close()
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
