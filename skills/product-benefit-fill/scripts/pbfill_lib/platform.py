# -*- coding: utf-8 -*-
"""Plan-book platform, public preview route.

What this module does: send a product's input configuration (pbData) to the benefit-illustration endpoint and
save the year-by-year table it returns. No account, token or cookie is used and no plan book is generated.

What it never does: log in, read or store a session, solve or avoid a captcha, call the signed
calculate/uploadResult endpoints, or retry after the platform says no. Any login / membership / verification /
frequency response writes a stop marker; further calls are refused until the user removes it.
"""
import copy
import datetime
import http.client
import json
import os
import re
import ssl
import time
import urllib.parse

HOST = 'bxs-planbook-api.winbaoxian.com'
DETAIL_PATH = '/planBook/interest/multipleDetailWithRemark'
RISK_CODES = ('002021', '002022', '002012', '002013', '002015')
RISK_WORDS = ('滑动验证', '次数已用完', '会员', '操作频繁', '操作过于频繁', '生成上限', '登录已失效', '重新登录', '请先登录', '未登录')
GAP_SECONDS = 1.5
DAILY_CAP = 200
STOP_FILE = 'PLATFORM_STOP.json'
LOG_FILE = 'platform_requests.jsonl'
PERSONAL = ('userInfo', 'agentInfo', 'card', 'user', 'userAppId', 'userOpenId', 'userUuid', 'mobile', 'phone', 'wechat',
            'avatar', 'headImg', 'nickName', 'realName', 'qrCode', 'token', 'appToken', 'wxToken')
PROD_DROP = ('hadChoice', 'hide', 'totalBaof', 'insShowArr', 'autoImport', 'restrictList', 'resultImport', 'adjustBaoeData',
             'companyId', 'companyName', 'enableInspectionTest', 'enablePurchase', 'feeValidateMode', 'handTag',
             'liabilityShowType', 'multipleInsured', 'required', 'salesChannelList', 'titleName', 'trade',
             'valueAddedServiceIdList', 'xinCodeList')
INS_DROP = ('isFollowing', 'isCheckbox', 'groupKey', 'eleOrderList', 'baseBaof', 'calcResultRules', 'calculatorType',
            'eleOrderInResult', 'multipleInsured', 'saleStatus')


class Stop(Exception):
    pass


def scrub(o):
    """Drop personal / account fields before anything is written to disk."""
    if isinstance(o, dict):
        return {k: scrub(v) for k, v in o.items() if k not in PERSONAL}
    if isinstance(o, list):
        return [scrub(x) for x in o]
    return o


def risky(resp):
    if isinstance(resp, dict) and resp.get('success'):
        return False
    s = json.dumps(resp, ensure_ascii=False)[:800] if resp is not None else ''
    return any('"%s"' % c in s for c in RISK_CODES) or any(w in s for w in RISK_WORDS)


def good(resp):
    return isinstance(resp, dict) and bool(resp.get('success')) and isinstance(resp.get('data'), dict) \
        and bool(resp['data'].get('productDataList'))


# ---------------------------------------------------------------- input configuration -> pbData (no network)
def _init_data(init):
    d = init.get('data') if isinstance(init.get('data'), dict) and 'initData' in init['data'] else init
    if 'initData' not in d:
        raise ValueError('not a product input configuration: expected data.initData.inputData')
    return d


def main_insurance(init):
    p = _init_data(init)['initData']['inputData']['productGroupList'][0]['prodOrderList'][0]
    ms = [m for m in p['mulInsOrderList'] if m.get('required')]
    return p, (ms[0] if ms else p['mulInsOrderList'][0])


def elements(m):
    return {e['key']: e for e in m.get('eleOrderList') or []}


def _pick(e, want):
    for o in e.get('opts') or []:
        v = o.get('val') if isinstance(o, dict) else o
        if str(v) == str(want):
            return v
    return None


def _opts(e):
    return [(o.get('val'), o.get('desc')) if isinstance(o, dict) else (o, None) for o in e.get('opts') or []]


def _verify(init, sets, age, sex):
    """Walk the platform's own option tree so that only combinations it offers for this age/sex are sent."""
    p, m = main_insurance(init)
    vd = (_init_data(init).get('verifyData') or {}).get(str(p['key']), {}).get(str(m['key']))
    if not vd:
        return True, sets, ''
    em = elements(m)
    blocks = [b for b in vd if b.get('sex') in (None, sex, str(sex))
              and (b.get('minAge') or 0) <= age <= (b.get('maxAge') if b.get('maxAge') is not None else 200)]
    if not blocks:
        return False, sets, 'age %d / sex %s is not offered for this entry' % (age, sex)
    last = ''
    for b in blocks:
        node = b.get('conditionData') or {}
        s2 = dict(sets)
        ok = True
        notes = []
        while node.get('key') and node.get(node['key']):
            k = node['key']
            ch = node[k]
            cur = s2.get(k, em.get(k, {}).get('value'))
            hit = [c for c in ch if str(c.get('value')) == str(cur)]
            if not hit:
                if k not in sets and len(ch) == 1:
                    hit = ch
                    v = _pick(em.get(k, {}), ch[0]['value'])
                    s2[k] = v if v is not None else ch[0]['value']
                    notes.append('%s set to the only offered value %s' % (k, s2[k]))
                else:
                    ok = False
                    last = '%s=%s is not offered (offered: %s)' % (k, cur, [c.get('value') for c in ch])
                    break
            node = hit[0]
        if ok:
            return True, s2, '; '.join(notes)
    return False, sets, last


def choose(init, pay, term, premium=100000, age=40, sex=1, overrides=None):
    """Map a workbook group (premium term in years, 1 = single premium; term text such as
    '终身' / '保至105周岁；60岁起年领' / '10年') to the platform's input element values.
    Returns (sets or None, notes)."""
    p, m = main_insurance(init)
    em = elements(m)
    sets, notes = {}, []
    parts = [x.strip() for x in str(term).split('；')]
    dur, rest = parts[0], '；'.join(parts[1:])
    if 'pPeriod' not in em:
        return None, ['entry has no premium-term element']
    want = 'a1000' if int(pay) == 1 else 'b%d' % int(pay)
    v = _pick(em['pPeriod'], want)
    if v is None:
        return None, ['premium term %s is not offered: %s' % (want, _opts(em['pPeriod']))]
    sets['pPeriod'] = v
    ov = [o[0] for o in _opts(em.get('iDuration', {}))]
    w = None
    mm = re.fullmatch(r'(\d+)年', dur)
    if dur in ('终身', '保至终身'):
        w = next((x for x in ('a999', 'a106', 'a105') if x in ov), None)
    elif mm:
        w = 'b' + mm.group(1)
    else:
        mm = re.match(r'保至(\d+)', dur)
        if mm:
            w = 'a' + mm.group(1)
    if w is None or w not in ov:
        if len(ov) == 1:
            notes.append('insurance period: only %s is offered; used for "%s"' % (ov[0], dur))
            w = ov[0]
        else:
            return None, ['insurance period "%s" (%s) is not offered: %s' % (dur, w, ov)]
    sets['iDuration'] = w
    mm = re.search(r'(\d+)岁起', rest)
    ra = int(mm.group(1)) if mm else None
    if 'rAge' in em:
        e = em['rAge']
        if e.get('opts'):
            if ra is not None:
                v = _pick(e, 'a%d' % ra)
                if v is None:
                    v = _pick(e, ra)
                if v is None:
                    return None, ['annuity start age %d is not offered: %s' % (ra, _opts(e))]
                sets['rAge'] = v
        elif ra is not None:
            sets['rAge'] = ra
    elif ra is not None:
        notes.append('entry has no start-age element; wanted %d' % ra)
    if 'obtainMode' in em:
        v = _pick(em['obtainMode'], 1)      # 1 = paid yearly
        if v is not None:
            sets['obtainMode'] = v
    if 'extra1' in em:
        e = em['extra1']
        lab = e.get('labelName') or ''
        if '保证给付' in lab or '保证领取' in lab:
            sets['extra1'] = _pick(e, 2 if '不含保证给付' in rest else 1)
        else:
            notes.append('element extra1 [%s] left at its default %s: check that it matches the group' % (lab, e.get('value')))
    if 'calMethod' in em and _pick(em['calMethod'], 'baof2baoe') is not None:
        sets['calMethod'] = 'baof2baoe'
        sets['baof'] = premium
        mode = 'premium'
    elif 'baof' in em and not em['baof'].get('isHide') and 'calMethod' not in em:
        sets['baof'] = premium
        mode = 'premium'
    elif 'num' in em:
        mode = 'units'
    elif 'baoe' in em:
        mode = 'amount'
    else:
        mode = 'unknown'
    if mode != 'premium':
        notes.append('entry is priced by %s, not by premium: the first reply shows the premium it implies; set %s so that '
                     'the yearly premium equals %s and fetch again' % (mode, 'num' if mode == 'units' else 'baoe', premium))
    for k, v in (overrides or {}).items():
        sets[k] = v
    ok, sets, note = _verify(init, sets, age, sex)
    if note:
        notes.append(note)
    if not ok:
        return None, notes
    notes.insert(0, 'mode=%s' % mode)
    return sets, notes


def build_pbdata(init, sets, age=40, sex=1):
    d = _init_data(init)
    p, m = main_insurance(init)
    ins = {}
    for k, v in m.items():
        if k == 'eleOrderList':
            for e in v:
                ins[e['key']] = sets[e['key']] if e['key'] in sets else e.get('value')
            continue
        if k in INS_DROP or k.endswith('Desc'):
            continue
        ins[k] = v
    for k, v in sets.items():
        ins.setdefault(k, v)
        ins[k] = v
    ins['title'] = True
    prod = {k: copy.deepcopy(v) for k, v in p.items() if k not in PROD_DROP and k != 'mulInsOrderList'}
    prod['mulInsOrderList'] = [ins]
    prod['common'] = {}
    person = lambda key, extra: dict({'title': True, 'name': '', 'sex': sex, 'age': age, 'birthday': None, 'key': key,
                                      'createArchive': False, 'cid': None}, **extra)
    return {'insuranceTypeId': d['pbId'],
            'personData': {'personOrderList': [person('insuredInfo', {}), person('applicantInfo', {'applicantAndInsuredSame': 0})]},
            'productGroupList': [{'key': 'allMainInsData', 'prodOrderList': [prod]}]}


# ---------------------------------------------------------------- the one network call
def detail_body(pbdata, take, annuity, bonus='mid', acc_rate=0.0175):
    p = pbdata['productGroupList'][0]['prodOrderList'][0]
    m = p['mulInsOrderList'][0]
    return {'uuid': '', 'pbData': json.dumps(pbdata, ensure_ascii=False, separators=(',', ':')),
            'productId': p['key'], 'insuranceId': m['key'], 'mainInsuranceId': None, 'insuranceInterestFlag': False,
            'drawInsuranceId': None, 'accountInsuranceId': None, 'bonus': {'value': bonus}, 'takeType': {'value': None},
            'bonusTake': {'value': take}, 'pensionTake': {'value': 'annuity_cash' if annuity else None}, 'rate': {'value': None},
            'addAndTakeData': None, 'addAndTakeFlag': False, 'drawData': None, 'duplicateDrawData': None, 'accRate': acc_rate,
            'drawFlag': False, 'endBonusRate': None, 'yearBonusRate': None, 'bonusRealRate': None, 'singleCalType': 2,
            'annuityAddYears': None, 'bonusAddYears': None}


class Client(object):
    def __init__(self, state_dir, base_url=None, gap=GAP_SECONDS, daily_cap=DAILY_CAP):
        self.state = os.path.abspath(state_dir)
        if not os.path.isdir(self.state):
            os.makedirs(self.state)
        u = urllib.parse.urlparse(base_url or 'https://' + HOST)
        self.scheme, self.netloc = u.scheme, u.netloc
        self.gap = max(float(gap), GAP_SECONDS) if base_url is None else float(gap)
        self.cap = daily_cap
        self.sent = 0
        self._last = 0.0

    def _today(self):
        f = os.path.join(self.state, LOG_FILE)
        day = datetime.date.today().isoformat()
        if not os.path.exists(f):
            return 0
        with open(f, encoding='utf8') as fh:
            return sum(1 for l in fh if l.startswith('{"time": "%s' % day))

    def check(self):
        f = os.path.join(self.state, STOP_FILE)
        if os.path.exists(f):
            with open(f, encoding='utf8') as fh:
                why = fh.read()[:300]
            raise Stop('the platform refused an earlier request (%s). Tell the user; do not retry, change entry point, '
                       'parameters or network to get around it. Only the user may delete %s.' % (why, f))
        if self._today() >= self.cap:
            raise Stop('daily request cap of %d reached' % self.cap)

    def _log(self, rec):
        with open(os.path.join(self.state, LOG_FILE), 'a', encoding='utf8') as fh:
            fh.write(json.dumps(rec, ensure_ascii=False) + '\n')

    def _stop(self, why):
        with open(os.path.join(self.state, STOP_FILE), 'w', encoding='utf8') as fh:
            json.dump({'time': datetime.datetime.now().isoformat(timespec='seconds'), 'reason': why}, fh, ensure_ascii=False)
        raise Stop(why)

    def post_detail(self, body, tag=''):
        """One request, one connection, no retry."""
        self.check()
        wait = self.gap - (time.time() - self._last)
        if wait > 0:
            time.sleep(wait)
        cls = http.client.HTTPSConnection if self.scheme == 'https' else http.client.HTTPConnection
        kw = {'context': ssl.create_default_context()} if self.scheme == 'https' else {}
        conn = cls(self.netloc, timeout=60, **kw)
        rec = {'time': datetime.datetime.now().isoformat(timespec='seconds'), 'path': DETAIL_PATH, 'tag': tag,
               'take': body.get('bonusTake', {}).get('value')}
        try:
            conn.request('POST', DETAIL_PATH, body=json.dumps(body).encode('utf8'),
                         headers={'Content-Type': 'application/json', 'User-Agent': 'Mozilla/5.0',
                                  'Origin': 'https://pbf.winbaoxian.com', 'Referer': 'https://pbf.winbaoxian.com/'})
            r = conn.getresponse()
            raw = r.read()
            status = r.status
        except (OSError, http.client.HTTPException) as e:
            self._last = time.time()
            self.sent += 1
            rec.update(error=type(e).__name__)
            self._log(rec)
            raise Stop('network error (%s); not retried automatically' % type(e).__name__)
        finally:
            conn.close()
        self._last = time.time()
        self.sent += 1
        rec['http'] = status
        if status != 200:
            self._log(rec)
            self._stop('HTTP %s from the illustration endpoint' % status)
        try:
            resp = json.loads(raw.decode('utf8'))
        except ValueError:
            self._log(rec)
            self._stop('reply is not JSON')
        rec.update(success=bool(isinstance(resp, dict) and resp.get('success')), code=resp.get('code') if isinstance(resp, dict) else None)
        self._log(rec)
        if risky(resp):
            self._stop('login / membership / verification / frequency reply: %s' % json.dumps(resp, ensure_ascii=False)[:200])
        return resp


def fetch(pbdata, takes, annuity, out_dir, state_dir, base_url=None, tag='', gap=GAP_SECONDS, daily_cap=DAILY_CAP):
    """Fetch one group. Files already present are not requested again."""
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    cl = Client(state_dir, base_url, gap, daily_cap)
    status = {'tag': tag, 'route': 'public preview (pbData), no login, no plan book generated',
              'date': datetime.datetime.now().isoformat(timespec='seconds'), 'details': {}}
    with open(os.path.join(out_dir, 'pbdata.json'), 'w', encoding='utf8') as fh:
        json.dump(pbdata, fh, ensure_ascii=False)
    for take in takes:
        name = 'detail_%s.json' % (take or 'default')
        f = os.path.join(out_dir, name)
        if os.path.exists(f):
            status['details'][take or 'default'] = 'cached'
            continue
        body = detail_body(pbdata, take, annuity)
        resp = cl.post_detail(body, tag)
        body['pbData'] = '<see pbdata.json>'
        with open(f, 'w', encoding='utf8') as fh:
            json.dump({'request': body, 'response': scrub(resp)}, fh, ensure_ascii=False)
        status['details'][take or 'default'] = 'ok' if good(resp) else 'no data: %s' % json.dumps(resp, ensure_ascii=False)[:200]
    status['requests_sent'] = cl.sent
    with open(os.path.join(out_dir, 'status.json'), 'w', encoding='utf8') as fh:
        json.dump(status, fh, ensure_ascii=False)
    return status


def columns(path):
    """{key: column} of a saved reply, or raise with the platform's message."""
    with open(path, encoding='utf8') as fh:
        d = json.load(fh)
    r = d.get('response', d)
    if not good(r):
        raise ValueError('reply holds no illustration: %s' % json.dumps(r, ensure_ascii=False)[:200])
    return {c['key']: c for c in r['data']['productDataList']}
