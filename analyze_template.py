# -*- coding: utf-8 -*-
"""
Analyze Word template structure
"""

from docx import Document

doc = Document('chamcong/Nguyen Van Cuong.docx')
t = doc.tables[0]
print('Table structure:')
print(f'Rows: {len(t.rows)}, Cols: {len(t.columns)}')

print('\nFirst 10 rows detailed:')
for i, row in enumerate(t.rows[:10]):
    cells_text = []
    for c in row.cells:
        text = c.text[:25].replace('\n', ' ')
        cells_text.append(text)
    print(f'Row {i}: {cells_text}')
