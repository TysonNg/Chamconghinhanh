"""Build a daily camera report from the same ordered photos as the gallery."""
import io
import re
import zipfile
from datetime import time

from docx import Document
from docx.package import ImageParts
from docx.opc.packuri import PackURI
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from PIL import Image, ImageOps


class _IndexedImageParts(ImageParts):
    """Index immutable report images instead of rehashing all previous images.

    python-docx also searches a list for each available image number. Both
    operations become expensive when a month contains hundreds of pictures.
    This collection belongs to just this report; no library globals are changed.
    """
    def __init__(self, existing=()):
        super().__init__()
        self._by_sha1 = {}
        self._used_numbers = set()
        self._next_number = 1
        for part in existing:
            self.append(part)

    def append(self, part):
        super().append(part)
        self._by_sha1[part.sha1] = part
        self._used_numbers.add(part.partname.idx)

    def _get_by_sha1(self, sha1):
        return self._by_sha1.get(sha1)

    def _add_image_part(self, image):
        part = super()._add_image_part(image)
        self._by_sha1[image.sha1] = part
        return part

    def _next_image_partname(self, ext):
        while self._next_number in self._used_numbers:
            self._next_number += 1
        number = self._next_number
        self._used_numbers.add(number)
        self._next_number += 1
        return PackURI(f'/word/media/image{number}.{ext}')


def _new_report_document():
    doc = Document()
    package = doc.part.package
    # image_parts is python-docx's cached, read-only lazyproperty.
    package.__dict__['image_parts'] = _IndexedImageParts(package.image_parts)
    return doc


def _photo_table(doc):
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for column in table.columns:
        column.width = Cm(3.6)
    margins = OxmlElement('w:tblCellMar')
    for side in ('top', 'left', 'bottom', 'right'):
        margin = OxmlElement(f'w:{side}')
        margin.set(qn('w:w'), '40' if side in ('left', 'right') else '0')
        margin.set(qn('w:type'), 'dxa')
        margins.append(margin)
    table._tbl.tblPr.append(margins)
    return table


def _append_daily_report(doc, project_name, day, photos):
    """Append one day using the same photo selection for daily/monthly exports."""
    candidates = {'morning': [], 'afternoon': []}
    for photo in photos:
        try:
            sent = time.fromisoformat(photo.get('send_time') or '')
        except (TypeError, ValueError):
            continue
        if sent.tzinfo is not None:
            continue
        shift = photo.get('shift')
        if ((shift == 'morning' and time(7) <= sent <= time(15)) or
                (shift == 'afternoon' and sent >= time(16))):
            candidates[shift].append((sent, str(photo['path']), photo))
    for group in candidates.values():
        group.sort(key=lambda item: (item[0], item[1]))
    # Each shift has its own 15-image allowance, ordered by send time.
    photos = [item[2] for item in (candidates['morning'][:15] +
                                  candidates['afternoon'][:15])]
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(1.5)
    section.left_margin = section.right_margin = Cm(1.5)
    normal = doc.styles['Normal']
    normal.font.name = 'Arial'
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.keep_with_next = False
    doc.add_paragraph('Tổng hợp ảnh camera theo ngày', 'Title')
    doc.add_paragraph(f'Dự án: {project_name}')
    doc.add_paragraph(f'Ngày: {day.strftime("%d/%m/%Y")} — Tổng số: {len(photos)} ảnh')

    groups = (
        ('Ca sáng · 06:00–18:00', {'morning'}),
        ('Ca chiều · 18:00–06:00', {'afternoon'}),
    )
    previous_has_photos = False
    for label, shifts in groups:
        selected = [p for p in photos if p['shift'] in shifts]
        if label == 'Ảnh khác' and not selected:
            continue
        heading = doc.add_paragraph(f'{label} ({len(selected)} ảnh)', 'Heading 1')
        heading.paragraph_format.page_break_before = previous_has_photos
        if not selected:
            doc.add_paragraph('Không có ảnh')
            previous_has_photos = False
            continue
        table = _photo_table(doc)
        for offset in range(0, len(selected), 5):
            # Four bounded-height rows fit inside A4 with the report heading.
            # Start a fresh table rather than letting Word choose print breaks.
            if offset and offset % 20 == 0:
                continuation = doc.add_paragraph(f'{label} (tiếp theo)', 'Heading 1')
                continuation.paragraph_format.page_break_before = True
                doc.add_paragraph(f'{project_name} — {day.strftime("%d/%m/%Y")}')
                table = _photo_table(doc)
            row = table.add_row()
            row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
            for cell, photo in zip(row.cells, selected[offset:offset + 5]):
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                paragraph = cell.paragraphs[0]
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # Do not chain table rows together: Word otherwise moves the
                # image grid away from the report header on the first page.
                paragraph.paragraph_format.keep_with_next = False
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(2)
                try:
                    with Image.open(photo['path']) as original:
                        image = ImageOps.exif_transpose(original)
                        image.load()
                        # Normalize unsupported formats and keep camera reports compact.
                        image = image.convert('RGB')
                        # The printed cell is only 3.45 x 4.5 cm. This still
                        # exceeds 300 dpi while avoiding huge monthly DOCX files.
                        image.thumbnail((450, 600))
                        stream = io.BytesIO()
                        image.save(stream, format='JPEG', quality=82, optimize=True)
                        stream.seek(0)
                        scale = min(3.45 / image.width, 4.5 / image.height)
                        paragraph.add_run().add_picture(
                            stream, width=Cm(image.width * scale), height=Cm(image.height * scale))
                except (OSError, ValueError, Image.DecompressionBombError) as exc:
                    raise ValueError(f'Không thể xuất ảnh "{photo["path"].name}": tệp hỏng hoặc không đọc được') from exc
        previous_has_photos = True

def _save_report(doc, project_name, period):
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', project_name).strip().rstrip('.')
    return output, f'Anh_{safe_name}_{period}.docx'


def export_daily_photo_word(project_name, day, photos):
    """Return an in-memory DOCX; photos contain path, shift and send_time."""
    doc = _new_report_document()
    _append_daily_report(doc, project_name, day, photos)
    return _save_report(doc, project_name, day.isoformat())


def export_monthly_photo_word(project_name, period, days, progress_callback=None):
    """Return a ZIP containing one daily DOCX per calendar day in a month folder."""
    if not re.fullmatch(r'\d{4}-\d{2}', period):
        raise ValueError('Tháng phải có định dạng YYYY-MM')
    output = io.BytesIO()
    # DOCX files are already compressed; storing them avoids redundant ZIP work.
    with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_STORED) as archive:
        for index, (day, photos) in enumerate(days):
            daily_output, filename = export_daily_photo_word(project_name, day, photos)
            archive.writestr(f'{period}/{filename}', daily_output.getvalue())
            if progress_callback:
                progress_callback(index + 1, len(days))
    output.seek(0)
    safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', project_name).strip().rstrip('.')
    return output, f'Anh_{safe_name}_{period}.zip'
