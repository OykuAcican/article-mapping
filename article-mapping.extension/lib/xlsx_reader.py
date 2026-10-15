# -*- coding: utf-8 -*-
"""Minimal, dependency-free .xlsx reader.

Works in pyRevit's IronPython 2.7 engine and in CPython 3, because it only
uses the standard library (zipfile + xml). Every cell value is returned as
text; numbers are formatted without float noise (0.30000000000000004 -> 0.3).
"""
import re
import zipfile
import xml.etree.ElementTree as ET

_M = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
_P = '{http://schemas.openxmlformats.org/package/2006/relationships}'
_CELL_RE = re.compile(r'^([A-Z]+)(\d+)$')


def _col_index(letters):
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def _fmt_number(text):
    try:
        f = float(text)
    except ValueError:
        return text
    if f == int(f) and abs(f) < 1e15:
        return str(int(f))
    return '%.10g' % f


class XlsxBook(object):
    """Reads every sheet of an .xlsx file into lists of text rows."""

    def __init__(self, path):
        z = zipfile.ZipFile(path)
        try:
            self._shared = self._read_shared(z)
            self._paths = self._read_sheet_paths(z)
            self._data = {}
            for name, part in self._paths:
                self._data[name] = self._read_sheet(z, part)
        finally:
            z.close()

    @property
    def sheet_names(self):
        return [name for name, _ in self._paths]

    def rows(self, sheet_name):
        return self._data.get(sheet_name, [])

    # -- internals ---------------------------------------------------------
    @staticmethod
    def _read_shared(z):
        if 'xl/sharedStrings.xml' not in z.namelist():
            return []
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        out = []
        for si in root.findall(_M + 'si'):
            parts = [t.text or '' for t in si.findall(_M + 't')]
            for run in si.findall(_M + 'r'):
                parts.extend(t.text or '' for t in run.findall(_M + 't'))
            out.append(''.join(parts))
        return out

    @staticmethod
    def _read_sheet_paths(z):
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        targets = {}
        for rel in rels.findall(_P + 'Relationship'):
            target = rel.get('Target')
            if target.startswith('/'):
                target = target[1:]
            elif not target.startswith('xl/'):
                target = 'xl/' + target
            targets[rel.get('Id')] = target
        out = []
        for sheet in wb.find(_M + 'sheets').findall(_M + 'sheet'):
            part = targets.get(sheet.get(_R + 'id'))
            if part:
                out.append((sheet.get('name'), part))
        return out

    def _cell_text(self, c):
        ctype = c.get('t')
        if ctype == 'inlineStr':
            node = c.find(_M + 'is')
            if node is None:
                return ''
            return ''.join(t.text or '' for t in node.iter(_M + 't'))
        v = c.find(_M + 'v')
        if v is None or v.text is None:
            return ''
        if ctype == 's':
            return self._shared[int(v.text)]
        if ctype == 'b':
            return 'TRUE' if v.text == '1' else 'FALSE'
        if ctype in ('str', 'e'):
            return v.text
        return _fmt_number(v.text)

    def _read_sheet(self, z, part):
        root = ET.fromstring(z.read(part))
        data = root.find(_M + 'sheetData')
        cells = {}
        max_row = max_col = -1
        if data is None:
            return []
        next_row = 0
        for row in data.findall(_M + 'row'):
            r_attr = row.get('r')
            r_idx = int(r_attr) - 1 if r_attr else next_row
            next_row = r_idx + 1
            next_col = 0
            for c in row.findall(_M + 'c'):
                ref = c.get('r')
                m = _CELL_RE.match(ref) if ref else None
                c_idx = _col_index(m.group(1)) if m else next_col
                next_col = c_idx + 1
                text = self._cell_text(c).strip()
                if text:
                    cells[(r_idx, c_idx)] = text
                    max_row = max(max_row, r_idx)
                    max_col = max(max_col, c_idx)
        table = [[''] * (max_col + 1) for _ in range(max_row + 1)]
        for (r, c), text in cells.items():
            table[r][c] = text
        return table
