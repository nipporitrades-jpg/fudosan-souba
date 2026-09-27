# -*- coding: utf-8 -*-
"""
国交省 不動産情報ライブラリAPI（XIT001）から、網のエリアの「取引価格」と「成約価格」を取り、
市区町村 × 種別 × 築年帯 の相場（㎡単価の中央値など）にまとめる。
住宅購入プロジェクト 司令塔アプリ用  v1.3（2026-09-27：--auto〈GitHub Actions で毎週自動〉を追加｜v1.2：--history〈23区の過去の推移〉｜v1.1：政令市を区ごとに取る・--only）

■ APIキーの扱い（いちばん大事）
  - キーを置くのは、ユーザーのパスワード管理アプリと、ユーザー自身の GitHub リポジトリの Secrets（REINFOLIB_API_KEY）だけ。
    Claude・チャット・DB・出力ファイル・このリポジトリのファイルには入れない（2026-09-27 ユーザーの選択。decisions/d-20260927-03）
  - 読み方：環境変数 REINFOLIB_API_KEY があればそれを使い、無ければ実行のたびに入力を求める
    （入力した文字は画面に出ない。パスワード管理アプリからコピーして貼り付け → Enter）。--auto のときは入力を求めない
  - このスクリプトはキーを画面にもファイルにも書き出さない

■ 使い方（Windows の PowerShell で、このファイルを置いたフォルダに移動してから）
  py reinfolib_souba.py                     # 去年と今年のデータ（初回はこれ）
  py reinfolib_souba.py --years 2024 2025 2026
  py reinfolib_souba.py --raw               # 取引1件ずつの明細（CSV）も手元に残す（Claudeに渡さなくてよい）
  py reinfolib_souba.py --only 川崎市 横浜市 相模原市   # 名前が一致する市区町村だけ（足りない所の取り直し用）
  py reinfolib_souba.py --history           # 23区の2006年〜いまの推移（市場タブの地図と折れ線用。15分ほど）
  所要時間：市区町村 約90 × 年数 のリクエストを1.5秒おきに送るので、2年分で5分ほど

■ 自動（GitHub Actions。.github/workflows/souba.yml が毎週土曜の早朝に実行する）
  python3 reinfolib_souba.py --auto --out data [--force]
    1. 最新の四半期（取引価格・成約価格）を千代田区の今年のデータで確かめる（1〜2回の問い合わせ）
    2. 前回（data/status.json）から新しい四半期が出た・前回が失敗した・120日以上たった・--force のときだけ、
       相場の表（網の市区町村・去年と今年）と23区の推移（2006年〜）を全部取り直す（25分ほど）
    3. data/souba_reinfolib_latest.json・data/souba_reinfolib_hist_latest.json・data/status.json に書く。
       途中で取れなかった所があれば前の版を残し、status.json に失敗を書いて終了コード1（GitHub から失敗のメールが届く）
    → Claude の週次スイープ（日曜4:00）が status.json を見て、新しい版だけを司令塔アプリのDBに入れる

■ 出力
  同じフォルダに souba_reinfolib_YYYYMMDD.json（集計だけ。キーは入っていない）
  → このファイルを Claude とのチャットに添付すると、DB（souba）に入れて全物件の相場比を計算する
  （--auto のときは上のとおり data/ の決まった名前に書く）

■ 出典（PDL1.0。表示が必須）
  国土交通省 不動産情報ライブラリ https://www.reinfolib.mlit.go.jp/
"""
import argparse, csv, datetime as dt, getpass, gzip, hashlib, json, os, re, sys, time
import urllib.error, urllib.parse, urllib.request

VERSION = '1.3'
BASE = 'https://www.reinfolib.mlit.go.jp/ex-api/external/'
CREDIT = ('このサービスは、国土交通省の不動産情報ライブラリのAPI機能を使用していますが、'
          '提供情報の最新性、正確性、完全性等が保証されたものではありません。')
WAIT = 1.5   # 秒。連続で送らない（APIマニュアル：間隔をあけ、連続実行しない）
JST = dt.timezone(dt.timedelta(hours=9))   # GitHub のサーバーは UTC なので、日付は日本時間で書く


def now_jst():
    return dt.datetime.now(JST)


def today_jst():
    return now_jst().date()

# 網のエリア（decisions/d-20260920-06）。東京都は島しょ部（コード13360以上）を除く全市区町村
NEAR = {
    '14': ['大和市'],
    '11': ['川口市', '戸田市', '蕨市', '和光市', '朝霞市', '新座市', '所沢市', '草加市', '八潮市', '三郷市'],
    '12': ['浦安市', '市川市', '松戸市'],
}
# 政令市は、市区町村の一覧（XIT002）に「市」でしか出ないが、取引価格（XIT001）は区のコードでしか返らない。
# そのため区のコード（全国地方公共団体コード）をここに持つ。横浜市は北部5区だけ（d-20260920-06）
WARDS = [
    ('14131', '川崎市川崎区'), ('14132', '川崎市幸区'), ('14133', '川崎市中原区'), ('14134', '川崎市高津区'),
    ('14135', '川崎市多摩区'), ('14136', '川崎市宮前区'), ('14137', '川崎市麻生区'),
    ('14117', '横浜市青葉区'), ('14118', '横浜市都筑区'), ('14109', '横浜市港北区'), ('14113', '横浜市緑区'), ('14101', '横浜市鶴見区'),
    ('14151', '相模原市緑区'), ('14152', '相模原市中央区'), ('14153', '相模原市南区'),
]
PREF = {'13': '東京都', '14': '神奈川県', '11': '埼玉県', '12': '千葉県'}
TYPE_MAP = {'中古マンション等': '中古マンション', '宅地(土地と建物)': '戸建（土地と建物）', '宅地(土地)': '土地'}
PRICE_CAT = {'不動産取引価格情報': '取引価格', '成約価格情報': '成約価格'}
# 相場を歪める取引は集計から外す（件数は meta に残す）
SKIP_REMARKS = ['関係者間', '調停', '競売', '隣地', '私道', '瑕疵', '他の権利', '負担付']
ERA = {'明治': 1867, '大正': 1911, '昭和': 1925, '平成': 1988, '令和': 2018}


def api_key(interactive=True):
    k = os.environ.get('REINFOLIB_API_KEY', '').strip()
    if not k and interactive:
        k = getpass.getpass('不動産情報ライブラリのAPIキーを貼り付けて Enter（画面には出ません）: ').strip()
    if not k:
        raise SystemExit('APIキーがありません（自動のときは GitHub の Settings → Secrets and variables → Actions に REINFOLIB_API_KEY を入れる）')
    return k


def get(path, params, key, tries=4):
    """JSONを返す。データが無い条件は None、何度やっても失敗したら 'ERR'。キーは表示しない"""
    url = BASE + path + '?' + urllib.parse.urlencode(params)
    for i in range(tries):
        req = urllib.request.Request(url, headers={
            'Ocp-Apim-Subscription-Key': key, 'Accept-Encoding': 'gzip', 'User-Agent': 'fudosan-personal/' + VERSION})
        try:
            with urllib.request.urlopen(req, timeout=90) as r:
                raw = r.read()
            if raw[:2] == b'\x1f\x8b':
                raw = gzip.decompress(raw)
            return json.loads(raw.decode('utf-8'))
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return None
            if e.code in (401, 403):
                raise SystemExit('APIキーが通りませんでした（HTTP %d）。キーが正しいか確かめてください'
                                 '（PC：コピーし直して貼る／GitHub：Settings → Secrets の REINFOLIB_API_KEY を入れ直す）' % e.code)
            if e.code == 429 or e.code >= 500:
                time.sleep(WAIT * (4 ** i))
                continue
            print('  ! HTTP %d（%s）' % (e.code, path), file=sys.stderr)
            return 'ERR'
        except (urllib.error.URLError, TimeoutError, OSError, ValueError):
            time.sleep(WAIT * (4 ** i))
    return 'ERR'


def rows_of(res):
    if res is None or res == 'ERR':
        return []
    if isinstance(res, dict):
        return res.get('data') or []
    return res if isinstance(res, list) else []


def cities(key):
    """XIT002 で市区町村の一覧を取り、網のエリアだけを残す"""
    out = []
    for area in ('13', '14', '11', '12'):
        lst = rows_of(get('XIT002', {'area': area}, key))
        time.sleep(WAIT)
        if not lst:
            raise SystemExit('市区町村の一覧（%s）が取れませんでした。時間をおいてもう一度実行してください' % PREF[area])
        if area == '13':
            picked = [c for c in lst if str(c.get('id', '')) < '13360' and str(c.get('id', '')) != '13100']
        else:
            picked = [c for c in lst if any(str(c.get('name', '')).startswith(n) for n in NEAR[area])]
            # 政令市の「市全体」と「区」が両方あれば、区だけを使う（二重に数えない）
            names = [str(c.get('name', '')) for c in picked]
            picked = [c for c in picked if not any(o != c.get('name') and o.startswith(str(c.get('name'))) for o in names)]
        if area != '13':
            miss = [n for n in NEAR[area] if not any(str(c.get('name', '')).startswith(n) for c in picked)]
            if miss:
                print('  ! 一覧に見つからない名前：%s（ファイル冒頭の NEAR を直すと取れる）' % '・'.join(miss), file=sys.stderr)
        for c in picked:
            out.append({'code': str(c['id']), 'name': str(c['name']), 'pref': PREF[area]})
        if area == '14':
            out += [{'code': code, 'name': name, 'pref': PREF['14']} for code, name in WARDS]
    return out


def num(s):
    try:
        return float(str(s).replace(',', '').strip())
    except (TypeError, ValueError):
        return None


def year_of(s):
    """'2005年'・'平成17年'・'戦前' などを西暦に。分からなければ None"""
    s = str(s or '').strip()
    m = re.match(r'^(\d{4})', s)
    if m:
        return int(m.group(1))
    m = re.match(r'^(明治|大正|昭和|平成|令和)(元|\d+)', s)
    if m:
        return ERA[m.group(1)] + (1 if m.group(2) == '元' else int(m.group(2)))
    return None


def period_of(s):
    m = re.match(r'^(\d{4})年第(\d)四半期', str(s or ''))
    return (int(m.group(1)), int(m.group(2))) if m else (None, None)


def band(age):
    if age is None:
        return '築年不明'
    if age <= 5:
        return '築5年以内'
    if age <= 10:
        return '築6〜10年'
    if age <= 20:
        return '築11〜20年'
    if age <= 30:
        return '築21〜30年'
    return '築31年以上'


def pct(v, p):
    if not v:
        return None
    v = sorted(v)
    k = (len(v) - 1) * p
    f = int(k)
    c = min(f + 1, len(v) - 1)
    return v[f] + (v[c] - v[f]) * (k - f)


def stats(items):
    u = [x['unit'] for x in items if x['unit'] is not None]
    p = [x['price'] for x in items if x['price'] is not None]
    r2 = lambda x: None if x is None else round(x, 2)
    r0 = lambda x: None if x is None else round(x)
    return {'count': len(items), 'unit_count': len(u),
            'unit_median_man': r2(pct(u, 0.5)), 'unit_avg_man': r2(sum(u) / len(u)) if u else None,
            'unit_p25_man': r2(pct(u, 0.25)), 'unit_p75_man': r2(pct(u, 0.75)),
            'price_median_man': r0(pct(p, 0.5)), 'price_avg_man': r0(sum(p) / len(p)) if p else None}


def parse(r, city):
    t = TYPE_MAP.get(r.get('Type'))
    if not t:
        return None, 'type'
    rem = str(r.get('Remarks') or '')
    if any(w in rem for w in SKIP_REMARKS):
        return None, 'remarks'
    y, q = period_of(r.get('Period'))
    price = num(r.get('TradePrice'))
    if price is None or y is None:
        return None, 'broken'
    area = num(r.get('TotalFloorArea')) if t == '戸建（土地と建物）' else num(r.get('Area'))
    built = year_of(r.get('BuildingYear'))
    age = None if (built is None or t == '土地') else max(0, y - built)
    muni = str(r.get('Municipality') or city['name'])
    return {'pref': city['pref'], 'city': muni, 'group': group_of(muni),
            'type': t, 'band': '全体' if t == '土地' else band(age), 'data_type': PRICE_CAT.get(r.get('PriceCategory'), '不明'),
            'y': y, 'q': q, 'price': price / 1e4, 'unit': (price / 1e4 / area) if area and area > 0 else None,
            'built': built}, None


def group_of(muni):
    """政令市の区 → 市全体の名前。横浜市は北部5区だけなので、そうと分かる名前にする"""
    g = re.sub(r'^(.+?市).+区$', r'\1', muni)
    return '横浜市（北部5区）' if g == '横浜市' else g


def aggregate(recs):
    def put(d, k, x):
        d.setdefault(k, []).append(x)
    by_band, by_q = {}, {}
    for x in recs:
        for c in {x['city'], x['group']}:          # 政令市は区ごと＋市全体の両方
            put(by_band, (x['pref'], c, x['type'], x['band'], x['data_type']), x)
            put(by_band, (x['pref'], c, x['type'], '全体', x['data_type']), x) if x['band'] != '全体' else None
            put(by_q, (x['pref'], c, x['type'], x['data_type'], x['y'], x['q']), x)
    summary = []
    for (pref, city, t, b, dtp), items in sorted(by_band.items()):
        ys = sorted({(i['y'], i['q']) for i in items})
        rec = {'pref': pref, 'city': city, 'type': t, 'age_band': b, 'data_type': dtp,
               'period': '%d年第%d四半期〜%d年第%d四半期' % (ys[0] + ys[-1])}
        rec.update(stats(items))
        rec['key'] = '｜'.join([city, t, b, dtp, rec['period']])
        summary.append(rec)
    trend = []
    for (pref, city, t, dtp, y, q), items in sorted(by_q.items()):
        rec = {'pref': pref, 'city': city, 'type': t, 'age_band': '全体', 'data_type': dtp, 'period': '%d年第%d四半期' % (y, q)}
        rec.update(stats(items))
        trend.append(rec)
    return summary, trend


# ---------- --history：23区の過去の推移 ----------
TOKYO23 = [('13101', '千代田区'), ('13102', '中央区'), ('13103', '港区'), ('13104', '新宿区'), ('13105', '文京区'), ('13106', '台東区'),
           ('13107', '墨田区'), ('13108', '江東区'), ('13109', '品川区'), ('13110', '目黒区'), ('13111', '大田区'), ('13112', '世田谷区'),
           ('13113', '渋谷区'), ('13114', '中野区'), ('13115', '杉並区'), ('13116', '豊島区'), ('13117', '北区'), ('13118', '荒川区'),
           ('13119', '板橋区'), ('13120', '練馬区'), ('13121', '足立区'), ('13122', '葛飾区'), ('13123', '江戸川区')]
HT = {'中古マンション': 'M', '戸建（土地と建物）': 'H'}
HD = {'取引価格': 'T', '成約価格': 'S'}
# 築年帯（そろえる対象）。5年ごとの細かい帯。築年不明はそろえない（件数にだけ入る）
HB = ['築0〜5年', '築6〜10年', '築11〜15年', '築16〜20年', '築21〜25年', '築26〜30年', '築31〜35年', '築36〜40年', '築41年以上']


def hband(y, built):
    if built is None:
        return len(HB)
    a = max(0, y - built)
    return min(8, a // 5 - (1 if a % 5 == 0 and a > 0 else 0)) if a > 5 else 0
MIN_WIN = 10        # 直近4四半期で値を出すのに必要な件数（築年の分かる取引）
MIN_WARDS = 0.8     # 23区全体を出すのに必要な、区の重みの合計
ITER = 6            # 築年帯の効果と時点の効果を交互に求める回数


def med(v):
    return pct(v, 0.5) if v else None


def fit_series(smp, w_all):
    """築年をそろえた値の元になる「時点の効果」を求める（中央値で求めるので外れ値に強い）。
    smp＝[(四半期, 築年帯, log㎡単価)]。log㎡単価 ＝ 時点の効果 ＋ 築年帯の効果 ＋ ばらつき と考え、
    築年帯の効果は全期間で1つ（全期間の築年帯の割合で重み付けした平均が0になるようにそろえる）。
    こうすると、ある時期に新しい物件がほとんど売れていなくても、残りの物件から時点の効果が出せる"""
    import math
    bands = {}
    for q, b, lu in smp:
        bands.setdefault(b, []).append((q, lu))
    ok = {b for b, v in bands.items() if len(v) >= 3}
    smp = [x for x in smp if x[1] in ok]
    beta = {b: 0.0 for b in ok}
    byq = {}
    for q, b, lu in smp:
        byq.setdefault(q, []).append((b, lu))
    alpha = {}
    for _ in range(ITER):
        alpha = {q: med([lu - beta[b] for b, lu in v]) for q, v in byq.items()}
        beta = {b: med([lu - alpha[q] for q, lu in bands[b]]) for b in ok}
        tw = sum(w_all.get(b, 0) for b in ok) or 1
        c = sum(w_all.get(b, 0) * beta[b] for b in ok) / tw
        beta = {b: v - c for b, v in beta.items()}
    return beta, byq


def history_rows(buckets, prices, nward, base, total=True):
    """直近4四半期（その四半期を含む）ごとに、区×種別×データの値を出す。
    adj＝築年をそろえた㎡単価（時点の効果。全期間の築年帯の割合の物件を想定した値）。raw＝そのままの中央値。
    pmed＝取引総額の中央値。n4＝4四半期の件数"""
    import math
    keys = {}
    for (wi, t, d, qi, b) in buckets:
        keys.setdefault((wi, t, d), set()).add(qi)
    for (wi, t, d, qi) in prices:
        keys.setdefault((wi, t, d), set()).add(qi)
    rows, adj_of, wtot = [], {}, {}
    for (wi, t, d), qs in sorted(keys.items()):
        lo, hi = min(qs), max(qs)
        smp = [(q, b, math.log(u)) for q in range(lo, hi + 1) for b in range(len(HB)) for u in buckets.get((wi, t, d, q, b), [])]
        cnt = {}
        for q, b, lu in smp:
            cnt[b] = cnt.get(b, 0) + 1
        tot = sum(cnt.values()) or 1
        w_all = {b: c / tot for b, c in cnt.items()}
        beta, byq = fit_series(smp, w_all) if smp else ({}, {})
        wtot[(wi, t, d)] = sum(len(prices.get((wi, t, d, q), [])) for q in range(lo, hi + 1))
        for k in range(lo + 3, hi + 1):
            win = range(k - 3, k + 1)
            allp = [p for q in win for p in prices.get((wi, t, d, q), [])]
            if not allp:
                continue
            allu = [u for q in win for b in range(len(HB) + 1) for u in buckets.get((wi, t, d, q, b), [])]
            adjv = [lu - beta[b] for q in win for b, lu in byq.get(q, [])]
            adj = math.exp(med(adjv)) if len(adjv) >= MIN_WIN else None
            adj_of[(wi, t, d, k)] = adj
            rows.append([wi, t, d, k, len(allp), None if adj is None else round(adj, 2),
                         None if not allu else round(med(allu), 2), round(med(allp))])
    # 23区全体（番号 nward）：区の値を、全期間の件数の割合で重み付けして合わせる（区の顔ぶれの違いで動かないように）
    for t in (('M', 'H') if total else ()):
        for d in ('T', 'S'):
            ws = {wi: wtot.get((wi, t, d), 0) for wi in range(nward)}
            tw = sum(ws.values())
            if not tw:
                continue
            ks = sorted({k for (wi, tt, dd, k) in adj_of if tt == t and dd == d})
            for k in ks:
                got = cov = 0.0
                for wi in range(nward):
                    a = adj_of.get((wi, t, d, k))
                    if a is not None:
                        got += ws[wi] / tw * a; cov += ws[wi] / tw
                win = range(k - 3, k + 1)
                allu = [u for wi in range(nward) for q in win for b in range(len(HB) + 1) for u in buckets.get((wi, t, d, q, b), [])]
                allp = [p for wi in range(nward) for q in win for p in prices.get((wi, t, d, q), [])]
                rows.append([nward, t, d, k, len(allp), round(got / cov, 2) if cov >= MIN_WARDS else None,
                             None if not allu else round(med(allu), 2), None if not allp else round(med(allp))])
    return rows


def history(a, key, fname=None):
    """23区の推移を取って書く。fname があればその名前で（--auto 用）。(出力の中身, パス) を返す"""
    t0 = time.time()
    this = today_jst().year
    print('23区の一覧を取得中…')
    lst = rows_of(get('XIT002', {'area': '13'}, key))
    time.sleep(WAIT)
    names = {str(c.get('id')): str(c.get('name')) for c in lst}
    wards = [{'code': c, 'name': names.get(c, n)} for c, n in TOKYO23]
    if a.only:
        wards = [w for w in wards if any(w['name'].startswith(o) for o in a.only)]
        if not wards:
            raise SystemExit('--only に当てはまる区がありません：' + '・'.join(a.only))
    years = list(range(a.year_from, this + 1))
    base = years[0]
    print('対象 %d区 × %d〜%d年（%d回。1回ごとに%.1f秒あけるので、%d分ほどかかります）'
          % (len(wards), years[0], years[-1], len(wards) * len(years), WAIT, round(len(wards) * len(years) * (WAIT + 0.6) / 60)))
    buckets, prices = {}, {}
    skipped, errors, empty, n_req, used = {'type': 0, 'remarks': 0, 'broken': 0}, [], 0, 0, 0
    for wi, w in enumerate(wards):
        got = 0
        for y in years:
            res = get('XIT001', {'year': y, 'city': w['code']}, key)
            n_req += 1
            time.sleep(WAIT)
            if res == 'ERR':
                errors.append('%s %d年' % (w['name'], y)); continue
            if res is None:
                empty += 1; continue
            for r in rows_of(res):
                x, why = parse(r, {'name': w['name'], 'pref': '東京都'})
                if why:
                    skipped[why] += 1; continue
                t, d = HT.get(x['type']), HD.get(x['data_type'])
                if not t or not d:
                    continue
                qi = (x['y'] - base) * 4 + x['q'] - 1
                if qi < 0:
                    continue
                prices.setdefault((wi, t, d, qi), []).append(x['price'])
                u = x['unit']
                if u is not None and 5 <= u <= 2000:          # 明らかな誤り（㎡の桁違いなど）は㎡単価に使わない
                    buckets.setdefault((wi, t, d, qi, hband(x['y'], x['built'])), []).append(u)
                got += 1
        used += got
        print('  [%2d/%d] %-6s %6d件（%d分経過）' % (wi + 1, len(wards), w['name'], got, round((time.time() - t0) / 60)))
    rows = history_rows(buckets, prices, len(wards), base, total=not a.only)
    today = today_jst().strftime('%Y%m%d')
    out = {'meta': {'kind': 'history', 'source': '国土交通省 不動産情報ライブラリ（XIT001 不動産価格〈取引価格・成約価格〉情報）',
                    'source_url': 'https://www.reinfolib.mlit.go.jp/', 'license': '公共データ利用規約（PDL1.0）', 'credit': CREDIT,
                    'script': 'reinfolib_souba.py v' + VERSION, 'fetched': today_jst().isoformat(),
                    'fetched_at': now_jst().isoformat(timespec='seconds'), 'base_year': base,
                    'years': [years[0], years[-1]], 'wards': wards + ([{'code': '13100', 'name': '23区全体'}] if not a.only else []),
                    'partial': bool(a.only), 'requests': n_req, 'no_data': empty, 'errors': errors, 'records_used': used, 'skipped': skipped,
                    'types': {'M': '中古マンション', 'H': '戸建（土地と建物）'}, 'data_types': {'T': '取引価格', 'S': '成約価格'},
                    'cols': ['ward', 'type', 'data_type', 'quarter', 'n4', 'adj_man', 'median_man', 'price_median_man'],
                    'notes': ['quarter＝(年−base_year)×4＋(四半期−1)。値はその四半期までの直近4四半期（1年）をまとめたもの',
                              'adj_man＝築年をそろえた㎡単価（万円/㎡）。log㎡単価＝時点の効果＋築年帯（5年ごと9つ）の効果 として中央値で交互に求め、'
                              '全期間の築年帯の割合の物件を想定した値にしたもの。売れた物件の築年が年々古くなっても動かない。4四半期で築年の分かる取引が%d件未満なら null' % MIN_WIN,
                              '23区全体の adj_man は、区の値を全期間の件数の割合で合わせた値（区の重みの合計が%.0f%%以上のときだけ）' % (MIN_WARDS * 100),
                              'median_man＝そのままの㎡単価の中央値。price_median_man＝取引総額の中央値（万円）。n4＝4四半期の件数',
                              '中古マンション＝取引総額÷専有面積、戸建＝取引総額÷延床面積（土地の値段を含む）。特殊な取引（関係者間・競売など）は除外']},
           'rows': rows}
    os.makedirs(a.out, exist_ok=True)
    path = os.path.join(a.out, fname or 'souba_reinfolib_hist_%s%s.json' % (today, '_part' if a.only else ''))
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
    print('\n完了（%d分）。行 %d・使った取引 %d 件・データ無し %d・失敗 %d'
          % (round((time.time() - t0) / 60), len(rows), used, empty, len(errors)))
    if errors:
        print('  ! 取れなかった年：' + '・'.join(errors[:10]) + ('…' if len(errors) > 10 else '') + '（もう一度実行すると取り直せます）')
    if not fname:
        print('このファイルをClaudeに渡してください：\n  ' + path)
    return out, path


def run_table(a, key, fname=None):
    """網の市区町村の相場の表を取って書く。fname があればその名前で（--auto 用）。(出力の中身, パス) を返す"""
    t0 = time.time()
    print('市区町村の一覧を取得中…')
    cl = cities(key)
    if a.only:
        cl = [c for c in cl if any(c['name'].startswith(o) for o in a.only)]
        if not cl:
            raise SystemExit('--only に当てはまる市区町村がありません：' + '・'.join(a.only))
    print('対象 %d 市区町村 × %s年' % (len(cl), '・'.join(map(str, a.years))))
    recs, skipped, errors, empty, n_req = [], {'type': 0, 'remarks': 0, 'broken': 0}, [], 0, 0
    raw_rows = []
    for i, c in enumerate(cl, 1):
        got = 0
        for y in a.years:
            res = get('XIT001', {'year': y, 'city': c['code']}, key)
            n_req += 1
            time.sleep(WAIT)
            if res == 'ERR':
                errors.append('%s %d年' % (c['name'], y))
                continue
            if res is None:
                empty += 1
                continue
            for r in rows_of(res):
                x, why = parse(r, c)
                if why:
                    skipped[why] += 1
                    continue
                recs.append(x)
                got += 1
                if a.raw:
                    raw_rows.append(r)
        print('  [%2d/%d] %-14s %5d件' % (i, len(cl), c['name'], got))
    summary, trend = aggregate(recs)
    today = today_jst().strftime('%Y%m%d')
    out = {'meta': {'source': '国土交通省 不動産情報ライブラリ（XIT001 不動産価格〈取引価格・成約価格〉情報）',
                    'source_url': 'https://www.reinfolib.mlit.go.jp/', 'license': '公共データ利用規約（PDL1.0）',
                    'credit': CREDIT, 'script': 'reinfolib_souba.py v' + VERSION, 'fetched': today_jst().isoformat(),
                    'fetched_at': now_jst().isoformat(timespec='seconds'),
                    'years': a.years, 'cities': [c['name'] for c in cl], 'partial': bool(a.only), 'requests': n_req, 'no_data': empty,
                    'errors': errors, 'records_used': len(recs), 'skipped': skipped,
                    'notes': ['取引価格はアンケートによる実際の取引価格、成約価格は不動産流通機構の成約情報（2021年〜）。売出（掲載）価格と混ぜない',
                              '中古マンションの㎡単価＝取引総額÷専有面積。戸建（土地と建物）＝取引総額÷延床面積（土地の値段を含む）。土地＝取引総額÷土地面積',
                              '築年帯は取引の年−建築年。関係者間取引・調停・競売・隣地・私道・瑕疵・他の権利や負担付きの取引は除外',
                              '政令市（川崎市・横浜市・相模原市）は区ごとの行と、市全体の行の両方がある']},
           'summary': summary, 'trend': trend}
    os.makedirs(a.out, exist_ok=True)
    path = os.path.join(a.out, fname or 'souba_reinfolib_%s%s.json' % (today, '_part' if a.only else ''))
    with open(path, 'w', encoding='utf-8') as f:
        if fname:
            json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
        else:
            json.dump(out, f, ensure_ascii=False, indent=1)
    if a.raw and raw_rows:
        rp = os.path.join(a.out, 'reinfolib_raw_%s.csv' % today)
        cols = sorted({k for r in raw_rows for k in r.keys()})
        with open(rp, 'w', encoding='utf-8-sig', newline='') as f:
            w = csv.DictWriter(f, fieldnames=cols)
            w.writeheader()
            w.writerows(raw_rows)
        print('明細：' + rp)
    print('\n完了（%d分）。集計 %d 行・推移 %d 行・使った取引 %d 件・データ無し %d・失敗 %d'
          % (round((time.time() - t0) / 60), len(summary), len(trend), len(recs), empty, len(errors)))
    if not fname:
        print('このファイルをClaudeとのチャットに添付してください：\n  ' + path)
    return out, path


# ---------- --auto：GitHub Actions で毎週 ----------
AUTO_FILES = {'table': 'souba_reinfolib_latest.json', 'hist': 'souba_reinfolib_hist_latest.json'}
REFRESH_DAYS = 120      # 新しい四半期が出なくても、これより古ければ取り直す（国交省が過去分を直すことがあるため）
PCODE = {'不動産取引価格情報': 'T', '成約価格情報': 'S'}


def latest_quarters(key):
    """いま公表されている最新の四半期を、取引価格（T）・成約価格（S）ごとに調べる。
    千代田区の今年（無ければ去年）のデータを見る。問い合わせは1〜2回。見つからなければ世田谷区で同じことをする"""
    this = today_jst().year
    got = {}
    for code in ('13101', '13112'):
        for y in (this, this - 1):
            res = get('XIT001', {'year': y, 'city': code}, key)
            time.sleep(WAIT)
            if res == 'ERR':
                raise RuntimeError('最新の四半期を確かめられませんでした（市区町村 %s・%d年）' % (code, y))
            for r in rows_of(res):
                yy, q = period_of(r.get('Period'))
                d = PCODE.get(r.get('PriceCategory'))
                if yy and d and (yy, q) > got.get(d, (0, 0)):
                    got[d] = (yy, q)
            if 'T' in got and 'S' in got:
                break
        if got:
            break
    if not got:
        raise RuntimeError('最新の四半期が見つかりませんでした（千代田区・世田谷区の今年と去年にデータがない）')
    return {d: '%d年第%d四半期' % got[d] for d in sorted(got)}


def sha256_of(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def run_part(part, a, key):
    """表（table）か推移（hist）を一時ファイルに取り、途中で取れなかった所が無ければ決まった名前に置き換える。
    取れなかった所があれば前の版を残す（半端なデータで上書きしない）"""
    final = os.path.join(a.out, AUTO_FILES[part])
    tmp = AUTO_FILES[part].replace('.json', '.tmp.json')
    if part == 'table':
        out, path = run_table(a, key, fname=tmp)
    else:
        out, path = history(a, key, fname=tmp)
    m = out['meta']
    errs = list(m.get('errors') or [])
    info = {'file': AUTO_FILES[part], 'fetched': m.get('fetched'), 'fetched_at': m.get('fetched_at'),
            'records_used': m.get('records_used'), 'requests': m.get('requests'), 'errors': errs[:30], 'n_errors': len(errs)}
    if errs:
        os.remove(path)
        info['ok'] = False
        info['kept_previous'] = os.path.exists(final)
        return info
    os.replace(path, final)
    info.update(ok=True, sha256=sha256_of(final), bytes=os.path.getsize(final))
    if part == 'hist':
        info['period'] = '%d年〜' % m.get('base_year')
    return info


def auto(a, key=None):
    """毎週の自動実行。新しい四半期が出たときだけ全部取り直す。結果と状態を status.json に書く。
    うまくいけば 0、どこかで失敗したら 1 を返す（GitHub がユーザーに失敗のメールを送る）。
    キーが無い・通らないときも status.json に書いてから終わる（Claude のスイープが気づけるように）"""
    os.makedirs(a.out, exist_ok=True)
    sp = os.path.join(a.out, 'status.json')
    try:
        prev = json.load(open(sp, encoding='utf-8'))
    except (OSError, ValueError):
        prev = {}
    force = bool(a.force) or os.environ.get('FORCE', '').strip().lower() == 'true'
    st = {'script': 'reinfolib_souba.py v' + VERSION, 'last_checked': now_jst().isoformat(timespec='seconds'),
          'ok': True, 'error': None, 'ran': [], 'reason': '', 'latest': prev.get('latest') or {},
          'table': prev.get('table') or {}, 'hist': prev.get('hist') or {},
          'note': 'APIキーはこのリポジトリのどのファイルにも入っていない（GitHub の Secrets だけ）。値は国交省の公開データを集計したもの',
          'credit': CREDIT, 'source_url': 'https://www.reinfolib.mlit.go.jp/', 'license': '公共データ利用規約（PDL1.0）'}
    try:
        key = key or api_key(interactive=False)
        latest = latest_quarters(key)
        st['latest'] = latest
        newq = latest != (prev.get('latest') or {})
        today = today_jst()
        why = {}
        for part in ('table', 'hist'):
            p = prev.get(part) or {}
            try:
                age = (today - dt.date.fromisoformat(p.get('fetched'))).days
            except (TypeError, ValueError):
                age = None
            if force:
                why[part] = '手動で全部取り直し'
            elif newq:
                why[part] = '新しい四半期（%s）' % '・'.join('%s %s' % ({'T': '取引', 'S': '成約'}[d], v) for d, v in sorted(latest.items()))
            elif not p.get('ok'):
                why[part] = '前回がまだ・失敗'
            elif age is None or age > REFRESH_DAYS:
                why[part] = '前回から%s日' % age
        st['ran'] = sorted(why)
        st['reason'] = '／'.join('%s：%s' % ({'table': '相場の表', 'hist': '23区の推移'}[k], v) for k, v in sorted(why.items())) or '新しい四半期なし（取り直さない）'
        print('最新の四半期：%s → %s' % (json.dumps(latest, ensure_ascii=False), st['reason']))
        for part in ('table', 'hist'):
            if part not in why:
                continue
            info = run_part(part, a, key)
            st[part] = info if info['ok'] else dict(prev.get(part) or {}, ok=False, last_error=info)
            if not info['ok']:
                st['ok'] = False
                st['error'] = '%s：%d回の問い合わせが失敗（前の版を残した）' % ({'table': '相場の表', 'hist': '23区の推移'}[part], info['n_errors'])
    except SystemExit as e:              # キーが通らない・一覧が取れない など
        st['ok'], st['error'] = False, str(e)
    except Exception as e:               # 想定外。中身（キー）は出さず、種類と文だけ
        st['ok'], st['error'] = False, '%s: %s' % (type(e).__name__, e)
    with open(sp, 'w', encoding='utf-8') as f:
        json.dump(st, f, ensure_ascii=False, indent=1)
    print('状態：%s（%s）' % ('OK' if st['ok'] else '失敗', st['error'] or st['reason']))
    return 0 if st['ok'] else 1


def main():
    ap = argparse.ArgumentParser(description='不動産情報ライブラリAPIから網のエリアの相場を集計する')
    this = today_jst().year
    ap.add_argument('--years', nargs='+', type=int, default=[this - 1, this])
    ap.add_argument('--out', default=os.path.dirname(os.path.abspath(__file__)))
    ap.add_argument('--raw', action='store_true', help='取引1件ずつの明細CSVも手元に保存する')
    ap.add_argument('--only', nargs='+', help='名前がこれで始まる市区町村だけを取る（例：--only 川崎市 横浜市）')
    ap.add_argument('--history', action='store_true', help='23区の過去の推移（--from の年〜今年）を取る。市場タブの地図と折れ線用')
    ap.add_argument('--from', dest='year_from', type=int, default=2006, help='--history で取り始める年（取引価格は2005年7月から、成約価格は2021年から）')
    ap.add_argument('--auto', action='store_true', help='毎週の自動実行（GitHub Actions）。新しい四半期が出たときだけ表と推移を取り直す')
    ap.add_argument('--force', action='store_true', help='--auto で、新しい四半期が無くても全部取り直す')
    a = ap.parse_args()
    if a.auto:
        if a.only or a.raw:
            raise SystemExit('--auto と --only・--raw は一緒に使えません')
        sys.exit(auto(a))
    key = api_key()
    if a.history:
        history(a, key)
        return
    run_table(a, key)


if __name__ == '__main__':
    main()
