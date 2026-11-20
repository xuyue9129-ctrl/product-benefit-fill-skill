# -*- coding: utf-8 -*-
"""Read and edit an .xlsx at the XML level.

Only the cells that are explicitly written change; every other part of the package
(styles, merged cells, column widths, other sheets, defined names) is copied byte for byte.
`Book.save` re-opens the result and proves that before it returns.
"""
import os
import re
import zipfile

from lxml import etree as E

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def q(tag):
    return '{%s}%s' % (NS, tag)


def col2n(s):
    n = 0
    for c in s:
        n = n * 26 + ord(c) - 64
    return n


def n2col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split(addr):
    m = re.match(r'^([A-Z]+)(\d+)$', addr)
    if not m:
        raise ValueError('bad cell address: %r' % (addr,))
    return m.group(1), int(m.group(2))


_REF = re.compile(r'(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])')


def shift_formula(text, drow, dcol):
    """Translate the relative references of a formula, as Excel does for a shared formula."""
    def repl(m):
        ca, col, ra, row = m.groups()
        if not ca:
            col = n2col(col2n(col) + dcol)
        if not ra:
            row = str(int(row) + drow)
        return ca + col + ra + row
    parts = text.split('"')
    for i in range(0, len(parts), 2):  # odd parts are string literals
        parts[i] = _REF.sub(repl, parts[i])
    return '"'.join(parts)


def numtxt(v):
    if isinstance(v, bool):
        raise ValueError('boolean is not a supported cell value')
    if isinstance(v, int):
        return str(v)
    v = float(v)
    if v != v or v in (float('inf'), float('-inf')):
        raise ValueError('NaN/inf cannot be written to a cell')
    return str(int(v)) if v.is_integer() and abs(v) < 1e15 else repr(v)


def _norm(e):
    return e.tag, tuple(sorted(e.attrib.items())), e.text, tuple(_norm(c) for c in e)


class Sheet(object):
    def __init__(self, book, name, part):
        self.book, self.name, self.part = book, name, part
        self.root = E.fromstring(book.data[part])
        self.sd = self.root.find(q('sheetData'))
        self.rows = {int(r.get('r')): r for r in self.sd}
        self._cells = {}
        self._shared = None
        self.changed = set()
        self.appended = []

    # ---------- reading ----------
    def rowcells(self, rn):
        d = self._cells.get(rn)
        if d is None:
            row = self.rows.get(rn)
            d = {} if row is None else {split(c.get('r'))[0]: c for c in row}
            self._cells[rn] = d
        return d

    def cell(self, addr):
        col, rn = split(addr)
        return self.rowcells(rn).get(col)

    def value(self, addr):
        """Cached value: float, str, or None when the cell is empty or was never calculated."""
        c = self.cell(addr)
        if c is None:
            return None
        t = c.get('t')
        if t == 'inlineStr':
            is_ = c.find(q('is'))
            return ''.join(is_.itertext()) if is_ is not None else ''
        v = c.find(q('v'))
        if v is None or v.text is None:
            return None
        if t == 's':
            return self.book.shared(int(v.text))
        if t in ('str', 'e'):
            return v.text
        if t == 'b':
            return v.text == '1'
        return float(v.text)

    def _shared_index(self):
        if self._shared is None:
            idx = {}
            for c in self.root.iter(q('c')):
                f = c.find(q('f'))
                if f is not None and f.get('t') == 'shared':
                    d = idx.setdefault(f.get('si'), {'master': None, 'cells': []})
                    d['cells'].append(c)
                    if f.text:
                        d['master'] = c
            self._shared = idx
        return self._shared

    def formula(self, addr):
        """Formula text without the leading '=', shared formulas resolved; None for a constant."""
        c = self.cell(addr)
        if c is None:
            return None
        f = c.find(q('f'))
        if f is None:
            return None
        if f.get('t') == 'shared' and not f.text:
            m = self._shared_index()[f.get('si')]['master']
            mc, mr = split(m.get('r'))
            col, rn = split(addr)
            return shift_formula(m.find(q('f')).text, rn - mr, col2n(col) - col2n(mc))
        return f.text or ''

    def last_row(self):
        return max(self.rows) if self.rows else 0

    # ---------- writing ----------
    def unshare(self, si):
        """Turn every member of a shared-formula group into an ordinary formula."""
        grp = self._shared_index().pop(si)
        m = grp['master']
        mc, mr = split(m.get('r'))
        text = m.find(q('f')).text
        for c in grp['cells']:
            col, rn = split(c.get('r'))
            f = c.find(q('f'))
            f.text = shift_formula(text, rn - mr, col2n(col) - col2n(mc))
            for k in ('t', 'si', 'ref'):
                f.attrib.pop(k, None)
            self.changed.add(c.get('r'))

    def _make(self, addr, kind, value, cache, style):
        c = E.Element(q('c'))
        c.set('r', addr)
        if style is not None:
            c.set('s', style)
        if kind == 'f':
            E.SubElement(c, q('f')).text = value.lstrip('=')
            if cache is not None:
                if isinstance(cache, str):
                    c.set('t', 'str')
                    E.SubElement(c, q('v')).text = cache
                else:
                    E.SubElement(c, q('v')).text = numtxt(cache)
        elif kind == 's':
            c.set('t', 'inlineStr')
            t = E.SubElement(E.SubElement(c, q('is')), q('t'))
            t.text = value
            t.set(XML_SPACE, 'preserve')
        elif kind == 'v':
            E.SubElement(c, q('v')).text = numtxt(value)
        else:
            raise ValueError('unknown cell kind %r' % (kind,))
        return c

    def style_of(self, addr):
        c = self.cell(addr)
        return None if c is None else c.get('s')

    def put(self, addr, kind, value, cache=None, style=None):
        """kind: 'v' number, 's' text, 'f' formula (with its calculated value in `cache`),
        'c' replace only the calculated value of an existing formula cell."""
        col, rn = split(addr)
        row = self.rows.get(rn)
        if row is None:
            raise KeyError('row %d does not exist in sheet %s' % (rn, self.name))
        cells = self.rowcells(rn)
        old = cells.get(col)
        if kind == 'c':
            if old is None or old.find(q('f')) is None:
                raise ValueError('%s is not a formula cell; cannot update its cached value' % addr)
            v = old.find(q('v'))
            if v is None:
                v = E.SubElement(old, q('v'))
            if isinstance(value, str):
                old.set('t', 'str')
                v.text = value
            else:
                old.attrib.pop('t', None)
                v.text = numtxt(value)
            self.changed.add(addr)
            return
        if old is not None:
            f0 = old.find(q('f'))
            if f0 is not None and f0.get('t') == 'shared':
                self.unshare(f0.get('si'))
            elif f0 is not None and f0.get('t') == 'array':
                raise ValueError('%s holds an array formula; refusing to replace it' % addr)
            if old.get('s') is not None:
                style = old.get('s')
        new = self._make(addr, kind, value, cache, style)
        if old is not None:
            row.replace(old, new)
        else:
            later = next((c for c in row if col2n(split(c.get('r'))[0]) > col2n(col)), None)
            if later is None:
                row.append(new)
            else:
                row.insert(row.index(later), new)
            sp = row.get('spans')
            if sp:
                a, b = sp.split(':')
                row.set('spans', '%d:%d' % (min(int(a), col2n(col)), max(int(b), col2n(col))))
        cells[col] = new
        self.changed.add(addr)

    def append_row(self, cells):
        """cells: [(column, kind, value, cache, style)] in any order. Returns the new row number."""
        rn = self.last_row() + 1
        row = E.SubElement(self.sd, q('row'))
        row.set('r', str(rn))
        cells = sorted(cells, key=lambda x: col2n(x[0]))
        row.set('spans', '%d:%d' % (col2n(cells[0][0]), col2n(cells[-1][0])))
        for col, kind, value, cache, style in cells:
            addr = '%s%d' % (col, rn)
            row.append(self._make(addr, kind, value, cache, style))
            self.changed.add(addr)
        self.rows[rn] = row
        self.appended.append(rn)
        return rn

    def _fix_dimension(self):
        dim = self.root.find(q('dimension'))
        if dim is None or not self.appended:
            return
        ref = dim.get('ref')
        m = re.match(r'^([A-Z]+)(\d+):([A-Z]+)(\d+)$', ref or '')
        if m and int(m.group(4)) < self.last_row():
            dim.set('ref', '%s%s:%s%d' % (m.group(1), m.group(2), m.group(3), self.last_row()))

    def tobytes(self):
        self._fix_dimension()
        return E.tostring(self.root, encoding='UTF-8', xml_declaration=True, standalone=True)


class Book(object):
    def __init__(self, path):
        self.path = os.path.abspath(path)
        with zipfile.ZipFile(self.path) as z:
            self.infos = z.infolist()
            self.data = {i.filename: z.read(i.filename) for i in self.infos}
        wb = E.fromstring(self.data['xl/workbook.xml'])
        rels = E.fromstring(self.data['xl/_rels/workbook.xml.rels'])
        target = {r.get('Id'): r.get('Target') for r in rels}
        self.sheets = {}
        for s in wb.iter(q('sheet')):
            t = target[s.get('{%s}id' % RNS)]
            self.sheets[s.get('name')] = t.lstrip('/') if t.startswith('/') else 'xl/' + t
        self._ss = None
        self._open = {}

    def shared(self, i):
        if self._ss is None:
            raw = self.data.get('xl/sharedStrings.xml')
            self._ss = [] if raw is None else [''.join(t.text or '' for t in si.iter(q('t')) if t.getparent().tag != q('rPh'))
                                                for si in E.fromstring(raw)]
        return self._ss[i]

    def find_sheet(self, *words):
        hits = [n for n in self.sheets if all(w in n for w in words)]
        if len(hits) != 1:
            raise KeyError('expected exactly one sheet containing %s, found %s' % ('+'.join(words), hits or 'none'))
        return hits[0]

    def sheet(self, name):
        if name not in self._open:
            self._open[name] = Sheet(self, name, self.sheets[name])
        return self._open[name]

    def locked(self):
        d, b = os.path.split(self.path)
        return os.path.exists(os.path.join(d, '~$' + b))

    def save(self, out):
        """Write a NEW file and prove that nothing except the written cells differs from the source."""
        out = os.path.abspath(out)
        if out == self.path:
            raise ValueError('refusing to overwrite the source workbook')
        if os.path.exists(out):
            raise ValueError('output already exists: %s' % out)
        d, b = os.path.split(out)
        if os.path.exists(os.path.join(d, '~$' + b)):
            raise ValueError('output workbook is open in Excel/WPS')
        new = dict(self.data)
        touched = {s.part: s for s in self._open.values() if s.changed}
        for part, s in touched.items():
            new[part] = s.tobytes()
        if d and not os.path.isdir(d):
            os.makedirs(d)
        tmp = out + '.tmp'
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as o:
            for i in self.infos:
                o.writestr(i, new[i.filename])
        try:
            report = self._verify(tmp, touched)
        except Exception:
            os.remove(tmp)
            raise
        os.replace(tmp, out)
        return report

    def _verify(self, out, touched):
        report = {}
        with zipfile.ZipFile(out) as b:
            if [i.filename for i in self.infos] != b.namelist():
                raise AssertionError('package parts differ')
            for i in self.infos:
                nm = i.filename
                if nm not in touched:
                    if self.data[nm] != b.read(nm):
                        raise AssertionError('untouched part changed: %s' % nm)
                    continue
                s = touched[nm]
                ar = E.fromstring(self.data[nm])
                br = E.fromstring(b.read(nm))
                ac = {c.get('r'): c for c in ar.iter(q('c'))}
                bc = {c.get('r'): c for c in br.iter(q('c'))}
                if not (set(ac) <= set(bc) and set(bc) - set(ac) <= s.changed):
                    raise AssertionError('%s: cells were lost or added outside the patch' % s.name)
                for k, c in ac.items():
                    if k not in s.changed and _norm(c) != _norm(bc[k]):
                        raise AssertionError('%s!%s changed but is not part of the patch' % (s.name, k))
                    if c.get('s') != bc[k].get('s'):
                        raise AssertionError('%s!%s lost its style' % (s.name, k))
                for x, y in zip(ar, br):
                    if x.tag == q('sheetData') or (x.tag == q('dimension') and s.appended):
                        continue
                    if _norm(x) != _norm(y):
                        raise AssertionError('%s: <%s> changed' % (s.name, x.tag.split('}')[1]))
                ra = [r.get('r') for r in ar.find(q('sheetData'))]
                rb = [r.get('r') for r in br.find(q('sheetData'))]
                if rb[:len(ra)] != ra or len(rb) - len(ra) != len(s.appended):
                    raise AssertionError('%s: existing rows were reordered or removed' % s.name)
                report[s.name] = {'cells_before': len(ac), 'cells_after': len(bc), 'changed': len(s.changed),
                                  'new_cells': len(set(bc) - set(ac)), 'appended_rows': len(s.appended)}
        return report
