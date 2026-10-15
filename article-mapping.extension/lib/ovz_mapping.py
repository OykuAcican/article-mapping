# -*- coding: utf-8 -*-
"""Shared mapping file + article list for the OVZ Keynote tool (no Revit imports).

Mapping file (.csv, ; separated, opens in Excel):
    Category;Family;Type;System;Size;Parameter;Value;Art_nr;Omschrijving;
    Updated by;Updated on;Project

One row = one rule "this combination of category/family/type/system/size gets
this value in this parameter". Art_nr/Omschrijving tell which article the value
came from (for OVZ_Keynote, Value = Art_nr). Files from the first version
(without Parameter/Value) are read as OVZ_Keynote rules. Several people can use the same file: on save the file is read
again and only *your* changes are merged in; rules that someone else changed
in the meantime are reported as conflicts instead of being overwritten.
"""
import codecs
import os
import shutil


def read_csv_rows(path):
    """Read a .csv saved by Excel (; , or tab separated, UTF-8 or ANSI)."""
    import csv
    text = None
    for enc in ('utf-8-sig', 'cp1252'):
        try:
            f = codecs.open(path, 'r', enc)
            try:
                text = f.read()
            finally:
                f.close()
            break
        except UnicodeDecodeError:
            continue
    lines = text.splitlines()
    if lines and lines[0].lower().startswith('sep='):
        lines = lines[1:]
    first = lines[0] if lines else ''
    delim = max([';', ',', '\t'], key=first.count)
    return [[c.strip() for c in row] for row in csv.reader(lines, delimiter=delim)]


def csv_line(values):
    out = []
    for v in values:
        v = u'{}'.format(v)
        if any(ch in v for ch in ';"\n\r'):
            v = '"' + v.replace('"', '""') + '"'
        out.append(v)
    return u';'.join(out)


FIELDS = ['Category', 'Family', 'Type', 'System', 'Size', 'Parameter', 'Value',
          'Art_nr', 'Omschrijving', 'Updated by', 'Updated on', 'Project']
DEFAULT_PARAM = 'OVZ_Keynote'


def make_key(category, family, type_name, system, size, parameter=DEFAULT_PARAM):
    parts = [parameter, category, family, type_name, system, size]
    return u'|'.join(u' '.join((p or u'').strip().lower().split()) for p in parts)


# --------------------------------------------------------------------------
# Article list ("alle delen")
# --------------------------------------------------------------------------
def load_articles(rows):
    """Return (list of (code, desc, unit, sub), {code: desc}, duplicate codes,
    column names, {code: {column: value}})."""
    hdr = None
    for i, row in enumerate(rows[:15]):
        if any(c.strip().lower() in ('art_nr.', 'art_nr') for c in row):
            hdr = i
            break
    if hdr is None:
        raise ValueError('No "Art_nr." column found in the first 15 rows.')
    low = [c.strip().lower() for c in rows[hdr]]

    def col(*names):
        for i, h in enumerate(low):
            if h in names:
                return i
        return -1
    ca, cd = col('art_nr.', 'art_nr'), col('vmsw_omschrijving', 'omschrijving')
    cu, cs = col('meeteenheid'), col('ovz_subdiscipline')

    def cell(row, i):
        return row[i].strip() if 0 <= i < len(row) else ''
    names = [c.strip() for c in rows[hdr]]
    columns = []
    for n in names:
        if n and n not in columns:
            columns.append(n)
    out, desc, dups, values = [], {}, [], {}
    for row in rows[hdr + 1:]:
        code = cell(row, ca)
        if not code:
            continue
        if code in desc:
            dups.append(code)
            continue
        desc[code] = cell(row, cd)
        out.append((code, cell(row, cd), cell(row, cu), cell(row, cs)))
        vals = {}
        for i, n in enumerate(names):
            if n and n not in vals:
                vals[n] = cell(row, i)
        values[code] = vals
    return out, desc, sorted(set(dups)), columns, values


# --------------------------------------------------------------------------
# Mapping file
# --------------------------------------------------------------------------
def load_mapping(path):
    """Return {key: record dict}. Missing file -> empty mapping."""
    if not path or not os.path.exists(path):
        return {}
    rows = read_csv_rows(path)
    if not rows:
        return {}
    low = [c.strip().lower() for c in rows[0]]
    idx = dict((f, low.index(f.lower())) for f in FIELDS if f.lower() in low)
    if 'Art_nr' not in idx and 'Value' not in idx:
        raise ValueError('The mapping file has no "Value" or "Art_nr" column.')
    out = {}
    for row in rows[1:]:
        rec = dict((f, row[i].strip() if i < len(row) else '') for f, i in idx.items())
        for f in FIELDS:
            rec.setdefault(f, '')
        if not rec['Parameter']:
            rec['Parameter'] = DEFAULT_PARAM
        if not rec['Value'] and rec['Parameter'] == DEFAULT_PARAM:
            rec['Value'] = rec['Art_nr']
        if not rec['Value']:
            continue
        key = make_key(rec['Category'], rec['Family'], rec['Type'],
                       rec['System'], rec['Size'], rec['Parameter'])
        out[key] = rec
    return out


def find_conflicts(disk, snapshot, changes):
    """Keys you changed that someone else also changed since you loaded the file.
    disk/snapshot: {key: record}; changes: {key: record or None (= delete)}."""
    out = []
    for key, mine in changes.items():
        theirs = disk.get(key)
        before = snapshot.get(key)
        theirs_v = theirs['Value'] if theirs else ''
        before_v = before['Value'] if before else ''
        mine_v = mine['Value'] if mine else ''
        if theirs_v != before_v and theirs_v != mine_v:
            out.append((key, theirs, mine))
    return out


def merge(disk, changes, skip_keys=()):
    merged = dict(disk)
    for key, rec in changes.items():
        if key in skip_keys:
            continue
        if rec is None:
            merged.pop(key, None)
        else:
            merged[key] = rec
    return merged


def write_mapping(path, mapping):
    """Write via a temp file and keep a .bak of the previous version."""
    def sort_key(rec):
        return (rec.get('Parameter', '').lower(), rec['Category'].lower(),
                rec['Family'].lower(), rec['Type'].lower(),
                rec['System'].lower(), rec['Size'].lower())
    lines = [csv_line(FIELDS)]
    for rec in sorted(mapping.values(), key=sort_key):
        lines.append(csv_line([rec.get(f, '') for f in FIELDS]))
    tmp = path + '.tmp'
    f = codecs.open(tmp, 'w', 'utf-8-sig')
    try:
        f.write(u'\r\n'.join(lines) + u'\r\n')
    finally:
        f.close()
    if os.path.exists(path):
        shutil.copyfile(path, path + '.bak')
        os.remove(path)
    os.rename(tmp, path)
