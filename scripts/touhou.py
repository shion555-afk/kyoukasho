# -*- coding: utf-8 -*-
"""
東翔高校 教科書購入票 自動生成ツール

【必要なファイル】
  - 前年（または当年）の購入票ファイル（例：2026東翔高校教科書購入票ver2.xlsm）
      → 新２年生・新３年生・新１年生 シートをマスター表として使用
  - 新2年・新3年の生データ（例：新2年次.xlsx／新3年次.xlsx）

【使い方】
  python touhou.py
"""
import re
import unicodedata
import openpyxl
import printlayout
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

FONT_NAME = 'Arial'
HEADER_FONT = Font(name=FONT_NAME, bold=True, size=10)
NORMAL_FONT = Font(name=FONT_NAME, size=10)
THIN = Side(style='thin', color='999999')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill('solid', fgColor='DDEBF7')


def style_header_row(ws, row, max_col):
    for c in range(1, max_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER


def book_price_tax(m):
    """(1冊ぶんの価格, 税額) を返す。

    注意：「価格」列（kakaku）は科目グループの小計で、グループ先頭の行にだけ入っている。
    1冊ぶんの金額は「単価」列（tanka）なので、必ずこちらを使うこと。
    課税品の税額は 単価 − 本体 で求める（元データに両方あるため逆算より正確）。
    """
    honntai = int(m['honntai']) if m['honntai'] else 0
    tanka = int(m['tanka']) if m['tanka'] else 0
    price = tanka if tanka else honntai
    if m['zei'] == '非':
        return price, 0
    tax = (price - honntai) if honntai else round(price - price / 1.1)
    return price, max(tax, 0)


# ============================================================
# 2年生
# ============================================================

ALIAS_2 = {
    '英語コミュニケーションII': 'コミュニケーション英語II',
    '英語課題探究I': '英語課題探求I',
    '簿記': '簿記(AB群で選択)',
    '会計ソフトウェア': '簿記(会計ソフトウェア)',
    '素描': '素描01',
}
COMBINED_2 = {unicodedata.normalize('NFKC', '物理基礎＋探究物理'): ['物理基礎', '探究物理']}
MATH2_SHORT = unicodedata.normalize('NFKC', '数学Ⅱ')
MATHB_NORM = unicodedata.normalize('NFKC', '数学Ｂ')


def norm_2(s):
    s = unicodedata.normalize('NFKC', s)
    return ALIAS_2.get(s, s)


def master_short_name_2(subject):
    parts = subject.split('\u3000')
    return norm_2(parts[-1] if len(parts) > 1 else subject)


def load_master_2nen(purchase_wb_path):
    wb = openpyxl.load_workbook(purchase_wb_path, data_only=True, keep_vba=True)
    ws = wb['新２年生']
    master = []
    for r in range(3, 84):
        subj = ws.cell(row=r, column=2).value
        if subj is None:
            continue
        master.append({
            'row': r, 'subject': subj, 'category': ws.cell(row=r, column=3).value,
            'publisher': ws.cell(row=r, column=4).value, 'code': ws.cell(row=r, column=5).value,
            'title': ws.cell(row=r, column=6).value, 'honntai': ws.cell(row=r, column=7).value,
            'tanka': ws.cell(row=r, column=8).value, 'kakaku': ws.cell(row=r, column=9).value,
            'zei': ws.cell(row=r, column=11).value,
        })
    return master


def load_students_2nen(raw_wb_path):
    wb = openpyxl.load_workbook(raw_wb_path, data_only=True)
    ws = wb['生徒データ']
    slot_names = ['選択数学', 'A', 'B', 'C', 'D', 'E', 'F']
    slot_cols = [6, 7, 8, 9, 10, 11, 12]
    NUMBER_FIX = {}  # 例: {('2', '須田\u3000陸斗'): 10}  ※データに誤記がある場合はここに追記
    students = []
    for r in range(4, ws.max_row + 1):
        kumi = ws.cell(row=r, column=2).value
        ban = ws.cell(row=r, column=3).value
        name = ws.cell(row=r, column=4).value
        keiretsu = ws.cell(row=r, column=5).value
        if kumi is None or name is None:
            continue
        fix_key = (str(kumi), name)
        if fix_key in NUMBER_FIX:
            ban = NUMBER_FIX[fix_key]
        slots = {}
        for sname, c in zip(slot_names, slot_cols):
            v = ws.cell(row=r, column=c).value
            if v not in (None, ''):
                slots[sname] = str(v).replace('○', '')
        if not slots:
            continue
        students.append({'組': kumi, '番': ban, '氏名': name, '系列': keiretsu, 'slots': slots})
    students.sort(key=lambda s: (s['組'], s['番']))
    return students


def is_checked_2nen(raw_norm_set, m):
    required = (m['category'] == '必修')
    short = master_short_name_2(m['subject'])
    if short in COMBINED_2:
        elective = all(part in raw_norm_set for part in COMBINED_2[short])
    else:
        elective = short in raw_norm_set
    if m['row'] in (12, 13) and short == MATH2_SHORT and MATHB_NORM in raw_norm_set:
        elective = False
    return required or elective


# ============================================================
# 3年生
# ============================================================

REQUIRED_ROWS_3 = {3, 4, 5}
ALIAS_3 = {
    'ソフトウエア活用': 'ソフトウェア活用',
    '数学課題探究': '数学課題探究B02',
    '原価計算': '原価計算(J群選択)',
    '英語コミュニケーションIII': 'コミュニケーション英語III',
}
PAIRED_ADDITIONS_3 = {'地学基礎02': ['地学基礎03']}


def norm_3(s):
    s = unicodedata.normalize('NFKC', s)
    return ALIAS_3.get(s, s)


def master_short_name_3(subject):
    parts = subject.split('\u3000')
    return norm_3(parts[-1] if len(parts) > 1 else subject)


def load_master_3nen(purchase_wb_path):
    wb = openpyxl.load_workbook(purchase_wb_path, data_only=True, keep_vba=True)
    ws = wb['新３年生']
    master = []
    for r in range(3, 100):
        subj = ws.cell(row=r, column=2).value
        if subj is None or '入力' in str(subj) or '印刷' in str(subj) or 'クリック' in str(subj):
            continue
        master.append({
            'row': r, 'subject': subj, 'publisher': ws.cell(row=r, column=4).value,
            'code': ws.cell(row=r, column=5).value, 'title': ws.cell(row=r, column=6).value,
            'honntai': ws.cell(row=r, column=7).value, 'tanka': ws.cell(row=r, column=8).value,
            'kakaku': ws.cell(row=r, column=9).value, 'zei': ws.cell(row=r, column=11).value,
        })
    return master


def load_students_3nen(raw_wb_path):
    wb = openpyxl.load_workbook(raw_wb_path, data_only=True)
    ws = wb['生徒データ']
    slot_names = ['G', 'H', 'H2', 'I1', 'I2', 'J1', 'J2', 'K', 'L']
    slot_cols = list(range(6, 15))
    NUMBER_FIX = {}  # 例: {('3', '阿部\u3000靖矢'): (4, 1)}
    students = []
    for r in range(4, ws.max_row + 1):
        kumi = ws.cell(row=r, column=2).value
        ban = ws.cell(row=r, column=3).value
        name = ws.cell(row=r, column=4).value
        keiretsu = ws.cell(row=r, column=5).value
        if kumi is None or name is None:
            continue
        fix_key = (str(kumi), name)
        if fix_key in NUMBER_FIX:
            kumi, ban = NUMBER_FIX[fix_key]
        slots = {}
        for sname, c in zip(slot_names, slot_cols):
            v = ws.cell(row=r, column=c).value
            if v not in (None, ''):
                slots[sname] = str(v).replace('○', '')
        if not slots:
            continue
        students.append({'組': kumi, '番': ban, '氏名': name, '系列': keiretsu, 'slots': slots})
    students.sort(key=lambda s: (s['組'], s['番']))
    return students


def matched_books_3(master, raw_norm_set, slot_val_norm=None, required_only=False):
    books = []
    for m in master:
        if required_only:
            if m['row'] in REQUIRED_ROWS_3:
                books.append(m)
            continue
        if m['row'] in REQUIRED_ROWS_3:
            continue
        short = master_short_name_3(m['subject'])
        if slot_val_norm is not None:
            if short == slot_val_norm:
                books.append(m)
        else:
            if short in raw_norm_set:
                books.append(m)
    return books


# ============================================================
# 1年生（共通リスト・名簿なし）
# ============================================================

def load_1nen(purchase_wb_path):
    wb = openpyxl.load_workbook(purchase_wb_path, data_only=True, keep_vba=True)
    ws = wb['新１年生']
    master = []
    for r in range(10, 33):
        title = ws.cell(row=r, column=6).value
        if not title:
            continue
        master.append({
            'publisher': ws.cell(row=r, column=3).value, 'code': ws.cell(row=r, column=4).value,
            'title': title, 'price': ws.cell(row=r, column=8).value,
        })
    geijutsu = []
    for r in [34, 35, 36]:
        geijutsu.append({
            'publisher': ws.cell(row=r, column=3).value, 'code': ws.cell(row=r, column=4).value,
            'title': ws.cell(row=r, column=6).value, 'price': ws.cell(row=r, column=8).value,
        })
    return master, geijutsu


# ============================================================
# 出力（共通処理・1行1冊形式）
# ============================================================

def build_workbook(master, students, out_path, grade_label, slot_labels, get_books_fn):
    """grade_label: '2' か '3'。slot_labels: [(表示名, slotキー), ...] （必修は自動で先頭に追加）"""
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = '名簿'
    ws1.append(['組', '番', '氏名', '系列'])
    style_header_row(ws1, 1, 4)
    for s in students:
        ws1.append([s['組'], s['番'], s['氏名'], s['系列']])
        for c in range(1, 5):
            ws1.cell(row=ws1.max_row, column=c).font = NORMAL_FONT
    ws1.column_dimensions['C'].width = 16

    ws2 = wb.create_sheet('集計')
    fixed_headers = ['教科・科目', '出版社', '教科書番号', '教科書・副教材名', '本体', '税込価格']
    ws2.append(fixed_headers + [f"{s['組']}組{s['番']}番 {s['氏名']}" for s in students])
    style_header_row(ws2, 1, len(fixed_headers) + len(students))
    ws2.freeze_panes = ws2.cell(row=2, column=len(fixed_headers) + 1)
    for m in master:
        row = [m.get('subject'), m['publisher'], m['code'], m['title'], m.get('honntai'), m.get('kakaku')]
        for s in students:
            raw_norm_set = set((norm_2(v) if grade_label == '2' else norm_3(v)) for v in s['slots'].values())
            if grade_label == '3':
                for base, extra in PAIRED_ADDITIONS_3.items():
                    if base in raw_norm_set:
                        raw_norm_set.update(extra)
            books = get_books_fn(s, m, raw_norm_set)
            row.append('○' if books else '')
        ws2.append(row)
        r = ws2.max_row
        for c in range(1, len(fixed_headers) + 1):
            ws2.cell(row=r, column=c).font = NORMAL_FONT
    ws2.column_dimensions['A'].width = 20
    ws2.column_dimensions['D'].width = 32

    ws3 = wb.create_sheet('購入明細')
    ws3.append(['組', '番', '氏名', '系列', '区分', '教科書番号', '発行所', '教科書・副教材名', '税区分', '税額', '価格'])
    style_header_row(ws3, 1, 11)
    for s in students:
        raw_norm_set = set((norm_2(v) if grade_label == '2' else norm_3(v)) for v in s['slots'].values())
        if grade_label == '3':
            for base, extra in PAIRED_ADDITIONS_3.items():
                if base in raw_norm_set:
                    raw_norm_set.update(extra)
        first = True
        hikazei_total = 0
        kazei_total = 0
        rows_to_write = build_student_rows(s, master, raw_norm_set, slot_labels, grade_label)
        for label, m in rows_to_write:
            price, tax = book_price_tax(m)
            if m['zei'] == '非':
                hikazei_total += price
            else:
                kazei_total += price
            ws3.append([
                s['組'] if first else None, s['番'] if first else None, s['氏名'] if first else None,
                s['系列'] if first else None, label, m['code'], m['publisher'], m['title'], m['zei'], tax, price,
            ])
            r = ws3.max_row
            for c in range(1, 12):
                ws3.cell(row=r, column=c).font = NORMAL_FONT
            first = False
        ws3.append([None, None, None, None, None, None, None, None, None, '非課税合計（教科書）', hikazei_total])
        ws3.append([None, None, None, None, None, None, None, None, None, '課税合計（準教科書）', kazei_total])
        ws3.append([None, None, None, None, None, None, None, None, None, '総合計', hikazei_total + kazei_total])
        for rr in range(ws3.max_row - 2, ws3.max_row + 1):
            ws3.cell(row=rr, column=10).font = HEADER_FONT
            ws3.cell(row=rr, column=11).font = HEADER_FONT
    ws3.column_dimensions['C'].width = 16
    ws3.column_dimensions['H'].width = 32

    wb.save(out_path)
    print(f"保存しました: {out_path}")
    print(f"生徒数: {len(students)}  教科書マスター行数: {len(master)}")


def build_student_rows(s, master, raw_norm_set, slot_labels, grade_label):
    rows = []
    if grade_label == '2':
        required = [m for m in master if m['category'] == '必修']
        rows += [('必修', m) for m in required]
        for label, slot_key in slot_labels:
            raw_val = s['slots'].get(slot_key)
            if not raw_val:
                continue
            raw_val_norm = norm_2(raw_val)
            for m in master:
                if m['category'] == '必修':
                    continue
                short = master_short_name_2(m['subject'])
                if short in COMBINED_2:
                    matched = all(part in raw_norm_set for part in COMBINED_2[short])
                else:
                    matched = (short == raw_val_norm)
                if matched and m['row'] in (12, 13) and short == MATH2_SHORT and MATHB_NORM in raw_norm_set:
                    matched = False
                if matched:
                    rows.append((label, m))
    else:
        required = matched_books_3(master, raw_norm_set, required_only=True)
        rows += [('必修', m) for m in required]
        for label, slot_key in slot_labels:
            raw_val = s['slots'].get(slot_key)
            if not raw_val:
                continue
            raw_val_norm = norm_3(raw_val)
            books = matched_books_3(master, raw_norm_set, slot_val_norm=raw_val_norm)
            for base, extra in PAIRED_ADDITIONS_3.items():
                if raw_val_norm == base:
                    for extra_name in extra:
                        books += [m for m in master if master_short_name_3(m['subject']) == extra_name]
            rows += [(label, m) for m in books]
    return rows


def get_books_fn_2(s, m, raw_norm_set):
    return [1] if is_checked_2nen(raw_norm_set, m) else []


def get_books_fn_3(s, m, raw_norm_set):
    if m['row'] in REQUIRED_ROWS_3:
        return [1]
    short = master_short_name_3(m['subject'])
    return [1] if short in raw_norm_set else []


def build_1nen_workbook(master, geijutsu, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '購入リスト'
    ws.append(['教科書番号', '発行所', '書名', '価格'])
    style_header_row(ws, 1, 4)
    total = 0
    for m in master:
        price = int(m['price'])
        total += price
        ws.append([m['code'], m['publisher'], m['title'], price])
        r = ws.max_row
        for c in range(1, 5):
            ws.cell(row=r, column=c).font = NORMAL_FONT
    ws.append([None, None, '合計', total])
    ws.cell(row=ws.max_row, column=3).font = HEADER_FONT
    ws.cell(row=ws.max_row, column=4).font = HEADER_FONT
    ws.append([])
    ws.append(['（参考）選択科目（芸術）※入学式後に代金引換で販売、合計には含めない'])
    ws.cell(row=ws.max_row, column=1).font = HEADER_FONT
    ws.append(['教科書番号', '発行所', '書名', '価格'])
    style_header_row(ws, ws.max_row, 4)
    for m in geijutsu:
        ws.append([m['code'], m['publisher'], m['title'], int(m['price'])])
        r = ws.max_row
        for c in range(1, 5):
            ws.cell(row=r, column=c).font = NORMAL_FONT
    ws.column_dimensions['C'].width = 40
    wb.save(out_path)
    print(f"保存しました: {out_path}")
    print(f"合計: {total}円")


# ============================================================
# メイン
# ============================================================

def build_1nen_print_layout(master, geijutsu, out_path, sale_date='', notes=None):
    """新1年：全生徒共通の1ページ印刷用リストを作る。"""
    books = [printlayout.make_book(m['code'], m['publisher'], m['title'], m['price']) for m in master]
    extras = [printlayout.make_book(g['code'], g['publisher'], g['title'], g['price']) for g in geijutsu]
    pages = [{'course': '全生徒共通', 'books': books, 'extras': extras,
              'extras_label': '（参考）選択科目（芸術）※入学式後に代金引換で販売、合計には含めません'}]
    printlayout.build_course_list_print(
        pages, out_path, '令和８年度教科書及び準教科書購入リスト', '北海道室蘭東翔高等学校１学年',
        sale_date=sale_date, notes=notes,
        subtitle='（定価には消費税に相当する金額が含まれています。）')


def build_print_layout(master, students, out_path, school_name, grade_label, slot_labels,
                        sale_date='', notes=None):
    """実際の購入票に近い「1生徒1ページ」の印刷用レイアウトを作る。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '購入票（印刷用）'

    TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14)
    SCHOOL_FONT = Font(name=FONT_NAME, bold=True, size=12)
    NAME_FONT = Font(name=FONT_NAME, bold=True, size=12)
    TOTAL_LABEL_FONT = Font(name=FONT_NAME, bold=True, size=11)
    TOTAL_VALUE_FONT = Font(name=FONT_NAME, bold=True, size=14)
    TABLE_HEADER_FONT = Font(name=FONT_NAME, bold=True, size=10)
    CENTER = Alignment(horizontal='center', vertical='center')
    RIGHT = Alignment(horizontal='right', vertical='center')

    col_widths = {'A': 3, 'B': 12, 'C': 10, 'D': 14, 'E': 42, 'F': 10, 'G': 8, 'H': 10}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    row = 1
    for s in students:
        raw_norm_set = set((norm_2(v) if grade_label == '2' else norm_3(v)) for v in s['slots'].values())
        if grade_label == '3':
            for base, extra in PAIRED_ADDITIONS_3.items():
                if base in raw_norm_set:
                    raw_norm_set.update(extra)
        rows_to_write = build_student_rows(s, master, raw_norm_set, slot_labels, grade_label)
        total = sum(book_price_tax(m)[0] for _, m in rows_to_write)

        ws.cell(row=row, column=2, value='令和８年度教科書及び準教科書購入票').font = TITLE_FONT
        row += 1
        ws.cell(row=row, column=2, value='（定価には消費税に相当する金額が含まれています。）').font = Font(name=FONT_NAME, size=9)
        row += 2

        ws.cell(row=row, column=2, value=school_name).font = SCHOOL_FONT
        ws.cell(row=row, column=5, value=f"{s['組']}組{s['番']}番　{s['氏名']}").font = NAME_FONT
        row += 1
        if sale_date:
            ws.cell(row=row, column=2, value=sale_date).font = Font(name=FONT_NAME, size=9)
        row += 1

        ws.cell(row=row, column=2, value='合計金額').font = TOTAL_LABEL_FONT
        ws.cell(row=row, column=3, value=total).font = TOTAL_VALUE_FONT
        ws.cell(row=row, column=3).alignment = RIGHT
        ws.cell(row=row, column=4, value='円').font = TOTAL_LABEL_FONT
        row += 2

        headers = ['区分', '教科書番号', '発行所', '書名', '税区分', '税額', '価格']
        for i, h in enumerate(headers):
            cell = ws.cell(row=row, column=2 + i, value=h)
            cell.font = TABLE_HEADER_FONT
            cell.alignment = CENTER
            cell.border = BORDER
            cell.fill = HEADER_FILL
        row += 1

        hikazei_total = 0
        kazei_total = 0
        for label, m in rows_to_write:
            price, tax = book_price_tax(m)
            if m['zei'] == '非':
                hikazei_total += price
            else:
                kazei_total += price
            values = [label, m['code'], m['publisher'], m['title'], m['zei'], tax, price]
            for i, v in enumerate(values):
                cell = ws.cell(row=row, column=2 + i, value=v)
                cell.font = NORMAL_FONT
                cell.border = BORDER
                if i in (5, 6):
                    cell.alignment = RIGHT
            row += 1

        for label, value in [
            ('非課税合計（教科書）', hikazei_total),
            ('課税合計（準教科書）', kazei_total),
            ('総合計', hikazei_total + kazei_total),
        ]:
            ws.cell(row=row, column=6, value=label).font = TABLE_HEADER_FONT
            cell = ws.cell(row=row, column=8, value=value)
            cell.font = TABLE_HEADER_FONT
            cell.alignment = RIGHT
            row += 1

        if notes:
            row += 1
            ws.cell(row=row, column=2, value='【お願い・注意事項】').font = Font(name=FONT_NAME, bold=True, size=9)
            row += 1
            for note in notes:
                ws.cell(row=row, column=2, value=f"・{note}").font = Font(name=FONT_NAME, size=9)
                row += 1

        row += 1
        ws.row_breaks.append(openpyxl.worksheet.pagebreak.Break(id=row - 1))

    ws.print_options.horizontalCentered = True
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    wb.save(out_path)
    print(f"保存しました: {out_path}")
    print(f"生徒数: {len(students)}")


def ask_and_build_print_layout(master, students, grade_label, slot_labels, school_name):
    ans = input("実際の購入票に近い印刷用レイアウトも作りますか？（y/n）: ").strip().lower()
    if ans != 'y':
        return
    print_out_path = input("印刷用ファイルの出力ファイル名: ").strip()
    sale_date = input("販売日の文言（空欄可）: ").strip()
    print("注意事項を1行ずつ入力してください（入力なしでEnterのみ押すと終了）:")
    notes = []
    while True:
        note = input(f"  注意事項{len(notes) + 1}: ").strip()
        if not note:
            break
        notes.append(note)
    build_print_layout(master, students, print_out_path, school_name, grade_label, slot_labels,
                        sale_date=sale_date, notes=notes)


def main():
    print("=== 東翔高校 教科書購入票 自動生成 ===")
    print("1: 新2年生")
    print("2: 新3年生")
    print("3: 新1年生（共通リスト）")
    grade = input("学年を選んでください（1/2/3）: ").strip()

    if grade == '1':
        purchase_path = input("購入票ファイルのパスを入力: ").strip()
        raw_path = input("2年生の生データ（新2年次.xlsx）のパスを入力: ").strip()
        out_path = input("出力ファイル名: ").strip()
        master = load_master_2nen(purchase_path)
        students = load_students_2nen(raw_path)
        slot_labels = [('選択数学', '選択数学'), ('A群', 'A'), ('B群', 'B'), ('C群', 'C'),
                       ('D群', 'D'), ('E群', 'E'), ('F群', 'F')]
        build_workbook(master, students, out_path, '2', slot_labels, get_books_fn_2)
        ask_and_build_print_layout(master, students, '2', slot_labels, '北海道室蘭東翔高等学校２学年')
    elif grade == '2':
        purchase_path = input("購入票ファイルのパスを入力: ").strip()
        raw_path = input("3年生の生データ（新3年次.xlsx）のパスを入力: ").strip()
        out_path = input("出力ファイル名: ").strip()
        master = load_master_3nen(purchase_path)
        students = load_students_3nen(raw_path)
        slot_labels = [('G群', 'G'), ('H群', 'H'), ('H②群', 'H2'), ('I①群', 'I1'), ('I②群', 'I2'),
                       ('J①群', 'J1'), ('J②群', 'J2'), ('K群', 'K'), ('L群', 'L')]
        build_workbook(master, students, out_path, '3', slot_labels, get_books_fn_3)
        ask_and_build_print_layout(master, students, '3', slot_labels, '北海道室蘭東翔高等学校３学年')
    elif grade == '3':
        purchase_path = input("購入票ファイルのパスを入力: ").strip()
        out_path = input("出力ファイル名: ").strip()
        master, geijutsu = load_1nen(purchase_path)
        build_1nen_workbook(master, geijutsu, out_path)
        opts = printlayout.ask_print_options('東翔高校_新1年_購入リスト（印刷用）.xlsx')
        if opts:
            build_1nen_print_layout(master, geijutsu, opts['path'],
                                    sale_date=opts['sale_date'], notes=opts['notes'])
    else:
        print("1・2・3のいずれかを入力してください。")


if __name__ == '__main__':
    main()
