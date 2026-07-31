# -*- coding: utf-8 -*-
"""
発注部数の照合（供給所の確注数 vs 現在の履修データから計算した必要数）

販売直前まで履修変更が入る学校向けに、
「いま何が何冊足りないか／余っているか」を出すためのモジュール。

不足＝至急手配が必要なので、レポートの一番上に赤字でまとめる。
"""
import os
import re
import sys
import unicodedata
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

FONT_NAME = 'Arial'
NORMAL_FONT = Font(name=FONT_NAME, size=10)
HEADER_FONT = Font(name=FONT_NAME, bold=True, size=10)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14)
SECTION_FONT = Font(name=FONT_NAME, bold=True, size=11)
SMALL_FONT = Font(name=FONT_NAME, size=9)
SHORT_FONT = Font(name=FONT_NAME, bold=True, size=10, color='C00000')

THIN = Side(style='thin', color='999999')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill('solid', fgColor='DDEBF7')
SHORT_FILL = PatternFill('solid', fgColor='FCE4E4')   # 不足＝薄い赤
OVER_FILL = PatternFill('solid', fgColor='FFF6DA')    # 余剰＝薄い黄
CENTER = Alignment(horizontal='center', vertical='center')
RIGHT = Alignment(horizontal='right', vertical='center')

# 供給所データの学校コード
SCHOOL_CODES = {
    '栄': 5532, '栄定時制': 5533, '清水丘': 5534, '工業': 5538,
    '東翔': 5540, '大谷': 5549, '海星': 5550, '大谷通信制': 5551,
}

# 書名の表記ゆれ（供給所 ⇔ こちら）。左右どちらの表記でも同じ本として扱う
TITLE_ALIASES = [
    ('高等学校標準古典探究', '高等学校標準古典探求'),
    ('高等学校標準現代の国語', '高等学校改訂版標準現代の国語'),
    ('保育基礎ようこそ，ともに育ち合う保育の世界へ', '保育基礎ようこそ、ともに育ち合う保育の世界へ'),
    ('新簿記新訂版', '新簿記改訂版'),
    ('改訂版物理基礎', '物理基礎'),
    ('高等学校改訂新化学基礎', '高等学校新化学基礎'),
]


def norm_title(s):
    """書名を比較用に正規化する（全角半角・空白・読点の違いを吸収）。"""
    if s is None:
        return ''
    s = unicodedata.normalize('NFKC', str(s))
    s = re.sub(r'[\s　、，,]+', '', s)
    return s.lower()


_ALIAS_MAP = {}
for a, b in TITLE_ALIASES:
    ka, kb = norm_title(a), norm_title(b)
    _ALIAS_MAP[ka] = ka
    _ALIAS_MAP[kb] = ka


def key_of(title):
    k = norm_title(title)
    return _ALIAS_MAP.get(k, k)


SUBJECT_COLS = ('科目名', '科目', '教科・科目', '科目名称')


def _pick_supply_sheet(wb):
    """複数シートがある場合は「科目名」列を持つシートを優先する。"""
    best = None
    for ws in wb.worksheets:
        hdr = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        if any(h in SUBJECT_COLS for h in hdr):
            return ws
        if best is None:
            best = ws
    return best


def load_supply(path, school_label):
    """供給所データから、指定した学校の教科書行を読み込む。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = _pick_supply_sheet(wb)
    ncol = ws.max_column
    hdr = [ws.cell(row=1, column=c).value for c in range(1, ncol + 1)]
    code = SCHOOL_CODES[school_label]
    rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ncol + 1)]
        d = dict(zip(hdr, vals))
        if d.get('学校名') is None or d.get('学校コード') != code:
            continue
        rows.append(d)
    return rows



# ============================================================
# 後期使用の教科書リスト（settings/後期使用の教科書.txt）
# ============================================================

SETTINGS_FILENAME = '後期使用の教科書.txt'


def _candidate_dirs():
    """設定ファイルを探す場所。exe化した場合も見つかるように複数見る。"""
    dirs = []
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        dirs += [os.path.join(here, '..', 'settings'), os.path.join(here, 'settings'), here]
    except NameError:
        pass
    exe = os.path.dirname(os.path.abspath(sys.executable))
    dirs += [os.path.join(exe, 'settings'), exe]
    cwd = os.getcwd()
    dirs += [os.path.join(cwd, 'settings'), os.path.join(cwd, '..', 'settings'), cwd]
    return dirs


def find_settings_file(filename=SETTINGS_FILENAME):
    for d in _candidate_dirs():
        path = os.path.join(d, filename)
        if os.path.exists(path):
            return os.path.normpath(path)
    return None


def load_later_term(school_label, path=None):
    """指定した学校の「後期使用」書名（正規化キー）の集合を返す。"""
    path = path or find_settings_file()
    keys = set()
    if not path:
        return keys
    try:
        with open(path, encoding='utf-8') as f:
            lines = f.readlines()
    except UnicodeDecodeError:
        with open(path, encoding='cp932') as f:
            lines = f.readlines()
    for line in lines:
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        parts = re.split(r'[,\t，]', line, maxsplit=1)
        if len(parts) != 2:
            continue
        school, title = parts[0].strip(), parts[1].strip()
        if school.replace('高校', '').replace('高等学校', '') == school_label and title:
            keys.add(key_of(title))
    return keys


class Needs:
    """学年ごとの必要数を集める入れ物。"""

    def __init__(self):
        self.books = {}          # key -> {'title','code','publisher'}
        self.by_grade = defaultdict(lambda: defaultdict(int))   # 学年 -> key -> 冊数
        self.grades = []         # 集計した学年（表示順）
        self.unknown_grades = [] # 人数が入力されず集計できなかった学年

    def add(self, grade, title, code, publisher, count):
        if not title or not count:
            return
        k = key_of(title)
        self.books.setdefault(k, {'title': title, 'code': code, 'publisher': publisher})
        self.by_grade[grade][k] += count
        if grade not in self.grades:
            self.grades.append(grade)

    def total(self, k):
        return sum(self.by_grade[g].get(k, 0) for g in self.grades)

    def note_unknown(self, grade):
        if grade not in self.unknown_grades:
            self.unknown_grades.append(grade)


def build_report(school_label, supply_rows, needs, out_path, later_keys=None):
    """照合レポートのExcelを作る。later_keys は後期使用として別枠にする書名キー。"""
    if later_keys is None:
        later_keys = load_later_term(school_label)
    sup_by_key = {}
    for d in supply_rows:
        sup_by_key.setdefault(key_of(d['書名']), d)

    shortages, overs, sames, only_mine, only_supply, later = [], [], [], [], [], []

    for k, meta in needs.books.items():
        need = needs.total(k)
        if k in later_keys:
            later.append({'meta': meta, 'need': need, 'key': k})
            continue
        if k in sup_by_key:
            d = sup_by_key[k]
            conf = d['確注生徒数'] or 0
            diff = need - conf
            row = {'meta': meta, 'sup': d, 'need': need, 'conf': conf, 'diff': diff, 'key': k}
            if diff > 0:
                shortages.append(row)
            elif diff < 0:
                overs.append(row)
            else:
                sames.append(row)
        else:
            only_mine.append({'meta': meta, 'need': need, 'key': k})

    used = set(needs.books)
    for k, d in sup_by_key.items():
        if k not in used and k not in later_keys:
            only_supply.append(d)

    shortages.sort(key=lambda x: -x['diff'])
    overs.sort(key=lambda x: x['diff'])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '発注部数の照合'
    widths = {'A': 3, 'B': 14, 'C': 12, 'D': 44, 'E': 9, 'F': 9, 'G': 9, 'H': 26}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    r = 1
    ws.cell(row=r, column=2, value=f'{school_label}高校　発注部数の照合').font = TITLE_FONT
    r += 1
    ws.cell(row=r, column=2,
            value='供給所の「確注生徒数」と、現在の履修データから計算した必要数の差').font = SMALL_FONT
    r += 1
    ws.cell(row=r, column=2,
            value=f"集計した学年：{'／'.join(needs.grades) if needs.grades else '（なし）'}").font = SMALL_FONT
    r += 1
    if needs.unknown_grades:
        ws.cell(row=r, column=2,
                value=f"※{'／'.join(needs.unknown_grades)}は人数が未入力のため必要数に含まれていません。"
                      f"その学年でも使う教科書は「不足」に見えることがあります").font = SMALL_FONT
        r += 1
    r += 1

    def table(title, rows, fill, is_short):
        nonlocal r
        ws.cell(row=r, column=2, value=title).font = SECTION_FONT
        r += 1
        heads = ['教番', '発行所', '書名', '確注数', '必要数', '過不足', '学年内訳']
        for i, h in enumerate(heads):
            c = ws.cell(row=r, column=2 + i, value=h)
            c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER; c.border = BORDER
        r += 1
        if not rows:
            ws.cell(row=r, column=2, value='（該当なし）').font = NORMAL_FONT
            r += 1
            return
        for x in rows:
            breakdown = '＋'.join(f"{g}{needs.by_grade[g][x['key']]}"
                                  for g in needs.grades if needs.by_grade[g].get(x['key']))
            vals = [x['sup']['教番'], x['sup']['発行者略称'], x['meta']['title'],
                    x['conf'], x['need'],
                    (f"＋{x['diff']}" if x['diff'] > 0 else str(x['diff']) if x['diff'] else '0'),
                    breakdown]
            for i, v in enumerate(vals):
                c = ws.cell(row=r, column=2 + i, value=v)
                c.font = SHORT_FONT if (is_short and i == 5) else NORMAL_FONT
                c.border = BORDER
                if fill:
                    c.fill = fill
                if i in (3, 4, 5):
                    c.alignment = RIGHT
            r += 1
        r += 1

    total_short = sum(x['diff'] for x in shortages)
    ws.cell(row=r, column=2, value=f'■ 不足（至急手配が必要）　{len(shortages)}点 / 合計{total_short}冊').font = SECTION_FONT
    ws.cell(row=r, column=2).fill = SHORT_FILL
    r += 1
    table('', shortages, SHORT_FILL, True)

    table(f'■ 余剰（返品・調整の候補）　{len(overs)}点', overs, OVER_FILL, False)
    table(f'■ 一致　{len(sames)}点', sames, None, False)

    if only_mine:
        ws.cell(row=r, column=2, value=f'■ 供給所データに無いもの　{len(only_mine)}点').font = SECTION_FONT
        r += 1
        ws.cell(row=r, column=2, value='準教科書（副教材）か、書名の表記が違う可能性があります').font = SMALL_FONT
        r += 1
        for i, h in enumerate(['教番', '発行所', '書名', '必要数']):
            c = ws.cell(row=r, column=2 + i, value=h)
            c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER; c.border = BORDER
        r += 1
        for x in sorted(only_mine, key=lambda y: -y['need']):
            for i, v in enumerate([x['meta']['code'], x['meta']['publisher'], x['meta']['title'], x['need']]):
                c = ws.cell(row=r, column=2 + i, value=v)
                c.font = NORMAL_FONT; c.border = BORDER
                if i == 3:
                    c.alignment = RIGHT
            r += 1
        r += 1

    if later:
        ws.cell(row=r, column=2, value=f'■ 後期使用（今回の照合では対象外）　{len(later)}点').font = SECTION_FONT
        r += 1
        ws.cell(row=r, column=2,
                value='settings／後期使用の教科書.txt に登録されているものです').font = SMALL_FONT
        r += 1
        for i, h in enumerate(['教番', '発行所', '書名', '後期の必要数']):
            c = ws.cell(row=r, column=2 + i, value=h)
            c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER; c.border = BORDER
        r += 1
        for x in sorted(later, key=lambda y: -y['need']):
            for i, v in enumerate([x['meta']['code'], x['meta']['publisher'],
                                    x['meta']['title'], x['need']]):
                c = ws.cell(row=r, column=2 + i, value=v)
                c.font = NORMAL_FONT; c.border = BORDER
                if i == 3:
                    c.alignment = RIGHT
            r += 1
        r += 1

    if only_supply:
        ws.cell(row=r, column=2,
                value=f'■ 確注はあるが今回の必要数に出てこないもの　{len(only_supply)}点').font = SECTION_FONT
        r += 1
        ws.cell(row=r, column=2,
                value='人数未入力の学年で使う教科書か、履修変更で不要になったもの（全量が返品候補）です').font = SMALL_FONT
        r += 1
        for i, h in enumerate(['教番', '発行所', '書名', '確注数']):
            c = ws.cell(row=r, column=2 + i, value=h)
            c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER; c.border = BORDER
        r += 1
        for d in sorted(only_supply, key=lambda y: -(y['確注生徒数'] or 0)):
            for i, v in enumerate([d['教番'], d['発行者略称'], d['書名'], d['確注生徒数']]):
                c = ws.cell(row=r, column=2 + i, value=v)
                c.font = NORMAL_FONT; c.border = BORDER
                if i == 3:
                    c.alignment = RIGHT
            r += 1

    ws.freeze_panes = ws.cell(row=1, column=4)
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(out_path)

    print(f'保存しました: {out_path}')
    print(f'  不足 {len(shortages)}点（合計{total_short}冊）／余剰 {len(overs)}点／'
          f'一致 {len(sames)}点／後期使用 {len(later)}点')
    if shortages:
        print('  ★至急手配:')
        for x in shortages[:10]:
            print(f"     {str(x['meta']['title'])[:36]:<36} 確注{x['conf']:>4} → 必要{x['need']:>4}  不足{x['diff']:+d}")
    return {'shortages': shortages, 'overs': overs, 'sames': sames, 'later': later}


# ============================================================
# 学校ごとの「必要数」の集め方
#   counts1 : 名簿が無い1年生などのコース別人数（辞書）。未入力なら None
# ============================================================

def needs_saka(purchase_path, raw2_path, raw3_path, counts1=None):
    """室蘭栄高校。counts1 = {'普通科': 人数, '理数科': 人数}"""
    import saka
    n = Needs()
    if raw2_path:
        m = saka.load_master_2nen(purchase_path)
        ss = saka.load_students_2nen(raw2_path)
        for b in m:
            c = sum(1 for s in ss if saka.is_checked_2nen(s, b))
            n.add('2年', b['title'], b['code'], b['publisher'], c)
    if raw3_path:
        m = saka.load_master_3nen(purchase_path)
        ss = saka.load_students_3nen(raw3_path)
        for b in m:
            c = sum(1 for s in ss if saka.is_checked_3nen(s, b))
            n.add('3年', b['title'], b['code'], b['publisher'], c)
    m1, g1 = saka.load_1nen(purchase_path)
    if counts1:
        for b in m1 + g1:
            c = 0
            if b.get('futsu') == '○':
                c += counts1.get('普通科', 0)
            if b.get('risu') == '○':
                c += counts1.get('理数科', 0)
            n.add('1年', b['title'], b['code'], b['publisher'], c)
    else:
        n.note_unknown('1年')
    return n


def needs_touhou(purchase_path, raw2_path, raw3_path, counts1=None):
    """東翔高校。counts1 = {'全生徒': 人数}"""
    import touhou
    n = Needs()
    for grade, mldr, sldr, path, getter in [
            ('2年', touhou.load_master_2nen, touhou.load_students_2nen, raw2_path, touhou.get_books_fn_2),
            ('3年', touhou.load_master_3nen, touhou.load_students_3nen, raw3_path, touhou.get_books_fn_3)]:
        if not path:
            continue
        m = mldr(purchase_path)
        ss = sldr(path)
        for b in m:
            c = 0
            for s in ss:
                rns = set((touhou.norm_2(v) if grade == '2年' else touhou.norm_3(v))
                          for v in s['slots'].values())
                if grade == '3年':
                    for base, extra in touhou.PAIRED_ADDITIONS_3.items():
                        if base in rns:
                            rns.update(extra)
                if getter(s, b, rns):
                    c += 1
            n.add(grade, b['title'], b['code'], b['publisher'], c)
    m1, g1 = touhou.load_1nen(purchase_path)
    if counts1:
        for b in m1 + g1:
            n.add('1年', b['title'], b['code'], b['publisher'], counts1.get('全生徒', 0))
    else:
        n.note_unknown('1年')
    return n


def _needs_flag_school(n, grade, master, students, code_key='code'):
    for i, b in enumerate(master):
        c = sum(1 for s in students if s['flags'][i] == 1)
        n.add(grade, b['title'], b.get(code_key), b.get('publisher'), c)


def needs_kaisei(wb_path, counts1=None):
    import kaisei
    n = Needs()
    for grade, ldr in [('2年', kaisei.load_2nen), ('3年', kaisei.load_3nen)]:
        m, ss, _ = ldr(wb_path)
        _needs_flag_school(n, grade, m, ss)
    if counts1:
        for b in kaisei.load_1nen(wb_path):
            n.add('1年', b['title'], b['code'], b['publisher'], counts1.get('全生徒', 0))
    else:
        n.note_unknown('1年')
    return n


def needs_otani(wb_path, counts1=None):
    """大谷高校。counts1 = {'難関進路': 人数, '文系・理系': 人数, '書Ⅰ': 人数, '高校美術': 人数}"""
    import otani
    n = Needs()
    _needs_flag_school(n, '2年',
                       otani.load_master(wb_path, '新2年購入表', 12, 25, otani.COLS_2NEN),
                       otani.load_students(wb_path, '新2年教科選択', 25, start_row=3))
    _needs_flag_school(n, '3年',
                       otani.load_master(wb_path, '新3年購入表', 16, 20, otani.COLS_3NEN),
                       otani.load_students(wb_path, '新3年教科選択', 20, start_row=3))
    m1, art = otani.load_1nen(wb_path)
    if counts1:
        for b in m1:
            c = 0
            if b.get('nankan') == '○':
                c += counts1.get('難関進路', 0)
            if b.get('bunkei_rikei') == '○':
                c += counts1.get('文系・理系', 0)
            n.add('1年', b['title'], b['code'], b['publisher'], c)
        for b in art:
            n.add('1年', b['title'], b['code'], b['publisher'], counts1.get(str(b['title']).strip(), 0))
    else:
        n.note_unknown('1年')
    return n


def needs_shimizuoka(path2, path3, path1=None, counts1=None):
    import shimizuoka
    n = Needs()
    if path2:
        m, ss, _ = shimizuoka.load_2nen(path2)
        _needs_flag_school(n, '2年', m, ss)
    if path3:
        m, ss, _ = shimizuoka.load_3nen(path3)
        _needs_flag_school(n, '3年', m, ss)
    if path1 and counts1:
        m1, g1 = shimizuoka.load_1nen(path1)
        for b in m1 + g1:
            n.add('1年', b['title'], b['code'], b['publisher'], counts1.get('全生徒', 0))
    else:
        n.note_unknown('1年')
    return n


def needs_kogyo(wb_path, counts=None):
    """工業高校。名簿が無いので全学年とも学科別の人数入力が必要。
    counts = {'1年': {'電気科': n, ...}, '2年': {...}, '3年': {...}}"""
    import kogyo
    n = Needs()
    counts = counts or {}
    plan = [('1年', kogyo.load_1nen, kogyo.columns_1nen),
            ('2年', kogyo.load_2nen, kogyo.columns_2nen),
            ('3年', kogyo.load_3nen, kogyo.columns_3nen)]
    for grade, loader, colfn in plan:
        c = counts.get(grade)
        if not c:
            n.note_unknown(grade)
            continue
        master = loader(wb_path)
        for name, matcher in colfn():
            num = c.get(name, 0)
            if not num:
                continue
            for b in master:
                if matcher(b):
                    n.add(grade, b['title'], b['code'], b['publisher'], num)
    return n
