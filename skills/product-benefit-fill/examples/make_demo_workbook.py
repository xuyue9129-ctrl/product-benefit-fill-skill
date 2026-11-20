# -*- coding: utf-8 -*-
"""Build a small synthetic workbook with the same layout as the real one (made-up products, made-up numbers).

Stored formula results are written as well, the way Excel saves them, so the tools can be tested without Excel.
"""
import os
import sys
import zipfile
from xml.sax.saxutils import escape

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'))

from pbfill_lib.calc import Block      # noqa: E402

RES = '产品利益信息——结果'
WORK = '产品利益信息——底稿'
PARAM, SUM_HEAD, BLOCK_HEAD = 2, 4, 9
PREMIUM = 100000.0


def _col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _n(col):
    n = 0
    for c in col:
        n = n * 26 + ord(c) - 64
    return n


class Grid(object):
    def __init__(self):
        self.rows = {}

    def put(self, addr, value=None, formula=None):
        col = addr.rstrip('0123456789')
        rn = int(addr[len(col):])
        self.rows.setdefault(rn, {})[col] = (value, formula)

    def xml(self):
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
        for rn in sorted(self.rows):
            out.append('<row r="%d">' % rn)
            for col in sorted(self.rows[rn], key=_n):
                v, f = self.rows[rn][col]
                a = '%s%d' % (col, rn)
                if f is not None:
                    t = ' t="str"' if isinstance(v, str) else ''
                    out.append('<c r="%s"%s><f>%s</f><v>%s</v></c>' % (a, t, escape(f), escape(str(v)) if isinstance(v, str) else repr(float(v))))
                elif isinstance(v, str):
                    out.append('<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (a, escape(v)))
                else:
                    out.append('<c r="%s"><v>%s</v></c>' % (a, repr(float(v))))
            out.append('</row>')
        out.append('</sheetData></worksheet>')
        return ''.join(out).encode('utf8')


def life_cv(n):
    return [0.0] + [round(PREMIUM * 1.02 ** y, 2) for y in range(1, n + 1)]


def build(path):
    res, work = Grid(), Grid()
    res.put('B1', '合成示例：产品与数字均为虚构，仅用于测试')
    for c, t in zip('BCDEFGHIJKLMNOPQ', ['序号', '产品', '利益类型', '保险期间', '交费期', '最高投保年龄', '保证IRR', '保证返本', '保证翻倍',
                                         '累生IRR', '累生返本', '累生翻倍', '增额IRR', '增额返本', '增额翻倍', '口径标记']):
        res.put('%s8' % c, t)
    work.put('A%d' % PARAM, '保费（元）')
    work.put('B%d' % PARAM, PREMIUM)
    work.put('C%d' % PARAM, '年龄')
    work.put('D%d' % PARAM, 40)
    work.put('E%d' % PARAM, '性别')
    work.put('F%d' % PARAM, '男')
    work.put('G%d' % PARAM, '累积利率')
    work.put('H%d' % PARAM, 0.0175)
    work.put('I%d' % PARAM, '演示利率')
    work.put('J%d' % PARAM, 0.035)
    work.put('A%d' % SUM_HEAD, '原序号')
    work.put('A%d' % BLOCK_HEAD, '组合号')
    work.put('C%d' % BLOCK_HEAD, '保单年度末')

    groups = [
        # result row, id, name, type, term, pay text, pay years, block data or None, dividend placeholder
        (9, 1, '示例终身寿险（分红型）', '分红', '终身', '趸交', 1, dict(E=life_cv(45), kind='H'), '待补演示'),
        (10, 2, '示例养老年金保险', '固收', '保至100岁；60岁起年领', '趸交', 1, None, '-'),
        (11, 3, '示例两全保险（分红型）', '分红', '终身', '3年', 3, dict(E=[0.0] + [round(min(y, 3) * PREMIUM * 0.8 * 1.025 ** max(0, y - 3), 2) for y in range(1, 46)],
                                                              kind='H', official={10: 20000.0, 20: 50000.0}), '待补演示'),
    ]
    nxt = BLOCK_HEAD + 1
    for k, (rr, pid, name, typ, term, pay_text, pay, data, placeholder) in enumerate(groups):
        sr = SUM_HEAD + 1 + k
        for c, v in zip('BCDEFG', [pid, name, typ, term, pay_text, 70]):
            res.put('%s%d' % (c, rr), v)
        work.put('A%d' % sr, pid)
        work.put('B%d' % sr, name)
        work.put('G%d' % sr, term)
        work.put('J%d' % sr, pay)
        work.put('K%d' % sr, 70)
        work.put('U%d' % sr, '合成示例。')
        work.put('AC%d' % sr, '合成示例来源。')
        vals = ['待补现价'] * 3 + [placeholder] * 6
        forms = [None] * 9
        if data:
            E = data['E']
            n = len(E) - 1
            blk = Block(PREMIUM, pay, E, [0.0] * (n + 1), [0.0] * (n + 1), E, None, data['kind'])
            g = blk.guarantee()
            a, t, z = nxt, nxt + blk.O, nxt + n
            miss = '%d岁内未达' % (40 + n)
            for i in range(n + 1):
                r = a + i
                work.put('A%d' % r, k + 1)
                work.put('B%d' % r, pid)
                work.put('C%d' % r, i)
                work.put('D%d' % r, blk.D[i], 'IF(C%d<$J$%d,$B$%d,0)' % (r, sr, PARAM))
                work.put('E%d' % r, E[i])
                work.put('F%d' % r, 0)
                work.put('G%d' % r, 0)
                work.put('H%d' % r, E[i])
                work.put('I%d' % r, 0.0, 'F%d' % r if i == 0 else 'I%d+F%d' % (r - 1, r))
                work.put('J%d' % r, blk.J[i], 'MIN(C%d,$J$%d)*$B$%d' % (r, sr, PARAM))
                work.put('K%d' % r, blk.K[i], 'IF(ISNUMBER(H{r}),H{r},"待补")'.format(r=r))
                work.put('L%d' % r, blk.L[i], 'IF(AND(C{r}=O{r},NOT(ISNUMBER(E{r}))),"待补",-D{r}+F{r}+G{r}+IF(C{r}=O{r},E{r},0))'.format(r=r))
                work.put('M%d' % r, g['f1'][i], 'IF(C{r}=0,0,IF(ISNUMBER(K{r}),IF(ROUND(K{r},6)>ROUND(J{r},6),1,0),0))'.format(r=r))
                work.put('N%d' % r, g['f2'][i], 'IF(C{r}=0,0,IF(ISNUMBER(K{r}),IF(ROUND(K{r},6)>ROUND(2*J{r},6),1,0),0))'.format(r=r))
                work.put('O%d' % r, blk.O)
                if i and i in data.get('official', {}):
                    work.put('P%d' % r, data['official'][i])
                work.put('AE%d' % r, name)
                work.put('AF%d' % r, term)
                work.put('AG%d' % r, pay_text)
                work.put('AH%d' % r, '结果表第%d行引用（底稿汇总行第%d行）' % (rr, sr))
            vals[0], forms[0] = g['irr'], 'IRR(L%d:L%d)' % (a, t)
            for j, (flag, got) in enumerate((('M', g['payback']), ('N', g['double'])), 1):
                if got is not None:
                    vals[j], forms[j] = got, 'INDEX(C{a}:C{z},MATCH(1,{c}{a}:{c}{z},0))'.format(a=a + 1, z=z, c=flag)
                else:
                    vals[j], forms[j] = miss, 'IF(COUNTIF({c}{a}:{c}{z},1)=0,"{m}","核查错误")'.format(a=a + 1, z=z, c=flag, m=miss)
            nxt = z + 1
        for j in range(9):
            sc, rc = 'LMNOPQRST'[j], 'HIJKLMNOP'[j]
            work.put('%s%d' % (sc, sr), vals[j], forms[j])
            res.put('%s%d' % (rc, rr), vals[j], '%s!%s%d' % (WORK, sc, sr))
        if placeholder != '-':
            res.put('Q%d' % rr, '此前标注示例')

    parts = {
        '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                               '<Default Extension="xml" ContentType="application/xml"/>'
                               '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                               '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                               '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                               '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
        '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
                           '<sheet name="%s" sheetId="1" r:id="rId1"/><sheet name="%s" sheetId="2" r:id="rId2"/></sheets></workbook>' % (RES, WORK),
        'xl/_rels/workbook.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                                      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
                                      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        'xl/styles.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                         '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill>'
                         '<fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
                         '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                         '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>',
    }
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, body in parts.items():
            z.writestr(name, body)
        z.writestr('xl/worksheets/sheet1.xml', res.xml())
        z.writestr('xl/worksheets/sheet2.xml', work.xml())
    return path


if __name__ == '__main__':
    print(build(sys.argv[1]))
