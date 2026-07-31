# -*- coding: utf-8 -*-
"""
高校教科書購入票 自動生成ツール（GUI版）

起動方法：
  python gui.py
  （実行ファイル化した場合は .exe をダブルクリック）

学校と学年を選ぶと、必要な入力ファイルの欄が自動で切り替わります。
「参照...」ボタンからExcelファイルを選んで、「実行」を押してください。
"""
import io
import os
import threading
import traceback
from contextlib import redirect_stdout

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import saka
import touhou
import kaisei
import otani
import shimizuoka
import kogyo
import shortage
import masterdiff


# ============================================================
# 各学校・学年ごとの処理
# ============================================================

def run_saka_students(paths, out_path, popts, grade, numbers=None):
    if grade == '2':
        master = saka.load_master_2nen(paths[0])
        students = saka.load_students_2nen(paths[1])
        checker = saka.is_checked_2nen
        school_name = '室蘭栄高等学校２学年'
    else:
        master = saka.load_master_3nen(paths[0])
        students = saka.load_students_3nen(paths[1])
        checker = saka.is_checked_3nen
        school_name = '室蘭栄高等学校３学年'
    saka.build_workbook(master, students, out_path, checker)
    if popts['enabled']:
        saka.build_print_layout(master, students, popts['path'], school_name, checker,
                                sale_date=popts['sale_date'], notes=popts['notes'])


def run_saka_1nen(paths, out_path, popts, grade, numbers=None):
    master, geijutsu = saka.load_1nen(paths[0])
    saka.build_1nen_workbook(master, geijutsu, out_path)
    if popts['enabled']:
        saka.build_1nen_print_layout(master, geijutsu, popts['path'],
                                     sale_date=popts['sale_date'], notes=popts['notes'])


TOUHOU_SLOTS_2 = [('選択数学', '選択数学'), ('A群', 'A'), ('B群', 'B'), ('C群', 'C'),
                  ('D群', 'D'), ('E群', 'E'), ('F群', 'F')]
TOUHOU_SLOTS_3 = [('G群', 'G'), ('H群', 'H'), ('H②群', 'H2'), ('I①群', 'I1'), ('I②群', 'I2'),
                  ('J①群', 'J1'), ('J②群', 'J2'), ('K群', 'K'), ('L群', 'L')]


def run_touhou_students(paths, out_path, popts, grade, numbers=None):
    if grade == '2':
        master = touhou.load_master_2nen(paths[0])
        students = touhou.load_students_2nen(paths[1])
        slots, getter = TOUHOU_SLOTS_2, touhou.get_books_fn_2
        school_name = '北海道室蘭東翔高等学校２学年'
    else:
        master = touhou.load_master_3nen(paths[0])
        students = touhou.load_students_3nen(paths[1])
        slots, getter = TOUHOU_SLOTS_3, touhou.get_books_fn_3
        school_name = '北海道室蘭東翔高等学校３学年'
    touhou.build_workbook(master, students, out_path, grade, slots, getter)
    if popts['enabled']:
        touhou.build_print_layout(master, students, popts['path'], school_name, grade, slots,
                                  sale_date=popts['sale_date'], notes=popts['notes'])


def run_touhou_1nen(paths, out_path, popts, grade, numbers=None):
    master, geijutsu = touhou.load_1nen(paths[0])
    touhou.build_1nen_workbook(master, geijutsu, out_path)
    if popts['enabled']:
        touhou.build_1nen_print_layout(master, geijutsu, popts['path'],
                                       sale_date=popts['sale_date'], notes=popts['notes'])


def run_kaisei_students(paths, out_path, popts, grade, numbers=None):
    if grade == '2':
        master, students, elective = kaisei.load_2nen(paths[0])
        school_name = '海星学院高等学校２学年'
    else:
        master, students, elective = kaisei.load_3nen(paths[0])
        school_name = '海星学院高等学校３学年'
    kaisei.build_workbook(master, students, elective, out_path)
    if popts['enabled']:
        kaisei.build_print_layout(master, students, elective, popts['path'], school_name,
                                  sale_date=popts['sale_date'], notes=popts['notes'])


def run_kaisei_1nen(paths, out_path, popts, grade, numbers=None):
    master = kaisei.load_1nen(paths[0])
    kaisei.build_1nen_workbook(master, out_path)
    if popts['enabled']:
        kaisei.build_1nen_print_layout(master, popts['path'],
                                       sale_date=popts['sale_date'], notes=popts['notes'])


def run_otani_students(paths, out_path, popts, grade, numbers=None):
    if grade == '2':
        master = otani.load_master(paths[0], '新2年購入表', 12, 25, otani.COLS_2NEN)
        students = otani.load_students(paths[0], '新2年教科選択', 25, start_row=3)
        school_name = '北海道大谷室蘭高等学校２学年'
    else:
        master = otani.load_master(paths[0], '新3年購入表', 16, 20, otani.COLS_3NEN)
        students = otani.load_students(paths[0], '新3年教科選択', 20, start_row=3)
        school_name = '北海道大谷室蘭高等学校３学年'
    otani.build_workbook(master, students, out_path)
    if popts['enabled']:
        otani.build_print_layout(master, students, popts['path'], school_name,
                                 sale_date=popts['sale_date'], notes=popts['notes'])


def run_otani_1nen(paths, out_path, popts, grade, numbers=None):
    master, art = otani.load_1nen(paths[0])
    otani.build_1nen_workbook(master, art, out_path)
    if popts['enabled']:
        otani.build_1nen_print_layout(master, art, popts['path'],
                                      sale_date=popts['sale_date'], notes=popts['notes'])


def run_shimizuoka_students(paths, out_path, popts, grade, numbers=None):
    if grade == '2':
        master, students, mark = shimizuoka.load_2nen(paths[0])
        school_name = '北海道室蘭清水丘高等学校２学年'
    else:
        master, students, mark = shimizuoka.load_3nen(paths[0])
        school_name = '北海道室蘭清水丘高等学校３学年'
    shimizuoka.build_workbook(master, students, mark, out_path)
    if popts['enabled']:
        shimizuoka.build_print_layout(master, students, mark, popts['path'], school_name,
                                      sale_date=popts['sale_date'], notes=popts['notes'])


def run_shimizuoka_1nen(paths, out_path, popts, grade, numbers=None):
    master, geijutsu = shimizuoka.load_1nen(paths[0])
    shimizuoka.build_1nen_workbook(master, geijutsu, out_path)
    if popts['enabled']:
        shimizuoka.build_1nen_print_layout(master, geijutsu, popts['path'],
                                           sale_date=popts['sale_date'], notes=popts['notes'])


def run_kogyo(paths, out_path, popts, grade, numbers=None):
    if grade == '1':
        master, columns = kogyo.build_1nen(paths[0], out_path)
    elif grade == '2':
        master, columns = kogyo.build_2nen(paths[0], out_path)
    else:
        master, columns = kogyo.build_3nen(paths[0], out_path)
    if popts['enabled']:
        kogyo.build_kogyo_print_layout(master, columns, popts['path'], grade,
                                       sale_date=popts['sale_date'], notes=popts['notes'])




# ---------- 発注部数の照合 ----------

def _report(school, paths, out_path, needs):
    sup = shortage.load_supply(paths[0], school)
    shortage.build_report(school, sup, needs, out_path)


def run_shortage_saka(paths, out_path, popts, grade, numbers=None):
    numbers = numbers or {}
    c1 = {k: numbers.get(k, 0) for k in ('普通科', '理数科')}
    n = shortage.needs_saka(paths[1], paths[2], paths[3], counts1=(c1 if any(c1.values()) else None))
    _report('栄', paths, out_path, n)


def run_shortage_touhou(paths, out_path, popts, grade, numbers=None):
    numbers = numbers or {}
    c1 = {'全生徒': numbers.get('全生徒', 0)}
    n = shortage.needs_touhou(paths[1], paths[2], paths[3], counts1=(c1 if c1['全生徒'] else None))
    _report('東翔', paths, out_path, n)


def run_shortage_kaisei(paths, out_path, popts, grade, numbers=None):
    numbers = numbers or {}
    c1 = {'全生徒': numbers.get('全生徒', 0)}
    n = shortage.needs_kaisei(paths[1], counts1=(c1 if c1['全生徒'] else None))
    _report('海星', paths, out_path, n)


def run_shortage_otani(paths, out_path, popts, grade, numbers=None):
    numbers = numbers or {}
    keys = ('難関進路', '文系・理系', '書 Ⅰ', '高校美術')
    c1 = {k: numbers.get(k, 0) for k in keys}
    n = shortage.needs_otani(paths[1], counts1=(c1 if any(c1.values()) else None))
    _report('大谷', paths, out_path, n)


def run_shortage_shimizuoka(paths, out_path, popts, grade, numbers=None):
    numbers = numbers or {}
    c1 = {'全生徒': numbers.get('全生徒', 0)}
    n = shortage.needs_shimizuoka(paths[1], paths[2], paths[3],
                                  counts1=(c1 if c1['全生徒'] else None))
    _report('清水丘', paths, out_path, n)


def run_shortage_kogyo(paths, out_path, popts, grade, numbers=None):
    numbers = numbers or {}
    counts = {
        '1年': {k: numbers.get('1年' + k, 0) for k in ('電気科', '建設科', '電子機械科')},
        '2年': {k: numbers.get('2年' + k, 0) for k in ('電気科', '建設科（土木）', '建設科（建築）', '電子機械科')},
        '3年': {k: numbers.get('3年' + k, 0) for k in
                ('環境土木科', '電気科（進学）', '電気科（専門）', '建築科',
                 '電子機械科（進学）', '電子機械科（専門）')},
    }
    counts = {g: c for g, c in counts.items() if any(c.values())}
    n = shortage.needs_kogyo(paths[1], counts=counts)
    _report('工業', paths, out_path, n)



# ---------- 教科書マスターの照合 ----------

def _mdiff(school, paths, out_path, mine_rows):
    sup = masterdiff.load_supply_all(paths[0], school)
    masterdiff.build_master_diff(school, sup, mine_rows, out_path)


def run_mdiff_saka(paths, out_path, popts, grade, numbers=None):
    _mdiff('栄', paths, out_path, masterdiff.mine_saka(paths[1]))


def run_mdiff_touhou(paths, out_path, popts, grade, numbers=None):
    _mdiff('東翔', paths, out_path, masterdiff.mine_touhou(paths[1]))


def run_mdiff_kaisei(paths, out_path, popts, grade, numbers=None):
    _mdiff('海星', paths, out_path, masterdiff.mine_kaisei(paths[1]))


def run_mdiff_otani(paths, out_path, popts, grade, numbers=None):
    _mdiff('大谷', paths, out_path, masterdiff.mine_otani(paths[1]))


def run_mdiff_shimizuoka(paths, out_path, popts, grade, numbers=None):
    _mdiff('清水丘', paths, out_path,
           masterdiff.mine_shimizuoka(paths[1], paths[2], paths[3]))


def run_mdiff_kogyo(paths, out_path, popts, grade, numbers=None):
    _mdiff('工業', paths, out_path, masterdiff.mine_kogyo(paths[1]))


# ============================================================
# 学校・学年ごとの設定
#   inputs   : 入力ファイル欄のラベル一覧
#   has_print: 印刷用レイアウト（1生徒1ページ）に対応しているか
# ============================================================

SCHOOLS = {
    '室蘭栄高校': {
        '教科書マスターの照合': {
            'inputs': ['供給所データ（需要数データ）', '購入票ファイル'],
            'has_print': False, 'run': run_mdiff_saka, 'grade': 'マスター照合',
            'default_out': 'マスター照合_栄高校.xlsx'},

        '発注部数の照合（全学年）': {
            'inputs': ['供給所データ（需要数データ）', '購入票ファイル（前年版など）',
                       '新2年生の生データ', '新3年生の生データ'],
            'numbers': [('1年 普通科の人数', '普通科'), ('1年 理数科の人数', '理数科')],
            'has_print': False, 'run': run_shortage_saka, 'grade': '照合',
            'default_out': '照合_栄高校_発注部数.xlsx'},

        '新1年（コース別共通リスト）': {
            'inputs': ['購入票ファイル（前年版など）'],
            'has_print': True, 'run': run_saka_1nen, 'grade': '1',
            'default_out': '栄高校_新1年_購入リスト.xlsx'},
        '新2年': {
            'inputs': ['購入票ファイル（前年版など）', '新2年生の生データ（科目選択一覧）'],
            'has_print': True, 'run': run_saka_students, 'grade': '2',
            'default_out': '栄高校_新2年_購入票.xlsx'},
        '新3年': {
            'inputs': ['購入票ファイル（前年版など）', '新3年生の生データ（科目選択一覧）'],
            'has_print': True, 'run': run_saka_students, 'grade': '3',
            'default_out': '栄高校_新3年_購入票.xlsx'},
    },
    '東翔高校': {
        '教科書マスターの照合': {
            'inputs': ['供給所データ（需要数データ）', '購入票ファイル'],
            'has_print': False, 'run': run_mdiff_touhou, 'grade': 'マスター照合',
            'default_out': 'マスター照合_東翔高校.xlsx'},

        '発注部数の照合（全学年）': {
            'inputs': ['供給所データ（需要数データ）', '購入票ファイル（前年版など）',
                       '新2年生の生データ（新2年次.xlsx）', '新3年生の生データ（新3年次.xlsx）'],
            'numbers': [('1年生の人数', '全生徒')],
            'has_print': False, 'run': run_shortage_touhou, 'grade': '照合',
            'default_out': '照合_東翔高校_発注部数.xlsx'},

        '新1年（共通リスト）': {
            'inputs': ['購入票ファイル（前年版など）'],
            'has_print': True, 'run': run_touhou_1nen, 'grade': '1',
            'default_out': '東翔高校_新1年_購入リスト.xlsx'},
        '新2年': {
            'inputs': ['購入票ファイル（前年版など）', '新2年生の生データ（新2年次.xlsx）'],
            'has_print': True, 'run': run_touhou_students, 'grade': '2',
            'default_out': '東翔高校_新2年_購入票.xlsx'},
        '新3年': {
            'inputs': ['購入票ファイル（前年版など）', '新3年生の生データ（新3年次.xlsx）'],
            'has_print': True, 'run': run_touhou_students, 'grade': '3',
            'default_out': '東翔高校_新3年_購入票.xlsx'},
    },
    '海星学院高校': {
        '教科書マスターの照合': {
            'inputs': ['供給所データ（需要数データ）', '教科書販売Excelファイル'],
            'has_print': False, 'run': run_mdiff_kaisei, 'grade': 'マスター照合',
            'default_out': 'マスター照合_海星学院高校.xlsx'},

        '発注部数の照合（全学年）': {
            'inputs': ['供給所データ（需要数データ）', '教科書販売Excelファイル'],
            'numbers': [('1年生の人数', '全生徒')],
            'has_print': False, 'run': run_shortage_kaisei, 'grade': '照合',
            'default_out': '照合_海星学院高校_発注部数.xlsx'},

        '新1年（共通リスト）': {
            'inputs': ['教科書販売Excelファイル'],
            'has_print': True, 'run': run_kaisei_1nen, 'grade': '1',
            'default_out': '海星学院高校_新1年_購入リスト.xlsx'},
        '新2年': {
            'inputs': ['教科書販売Excelファイル'],
            'has_print': True, 'run': run_kaisei_students, 'grade': '2',
            'default_out': '海星学院高校_新2年_購入票.xlsx'},
        '新3年': {
            'inputs': ['教科書販売Excelファイル'],
            'has_print': True, 'run': run_kaisei_students, 'grade': '3',
            'default_out': '海星学院高校_新3年_購入票.xlsx'},
    },
    '大谷高校（全日制）': {
        '教科書マスターの照合': {
            'inputs': ['供給所データ（需要数データ）', '教科選択・購入表Excelファイル'],
            'has_print': False, 'run': run_mdiff_otani, 'grade': 'マスター照合',
            'default_out': 'マスター照合_大谷高校.xlsx'},

        '発注部数の照合（全学年）': {
            'inputs': ['供給所データ（需要数データ）', '教科選択・購入表Excelファイル'],
            'numbers': [('1年 難関進路の人数', '難関進路'), ('1年 文系・理系の人数', '文系・理系'),
                        ('1年 芸術で書Ⅰを選ぶ人数', '書 Ⅰ'), ('1年 芸術で高校美術を選ぶ人数', '高校美術')],
            'has_print': False, 'run': run_shortage_otani, 'grade': '照合',
            'default_out': '照合_大谷高校_発注部数.xlsx'},

        '新1年（共通リスト）': {
            'inputs': ['教科選択・購入表Excelファイル'],
            'has_print': True, 'run': run_otani_1nen, 'grade': '1',
            'default_out': '大谷高校_新1年_購入リスト.xlsx'},
        '新2年': {
            'inputs': ['教科選択・購入表Excelファイル'],
            'has_print': True, 'run': run_otani_students, 'grade': '2',
            'default_out': '大谷高校_新2年_購入票.xlsx'},
        '新3年': {
            'inputs': ['教科選択・購入表Excelファイル'],
            'has_print': True, 'run': run_otani_students, 'grade': '3',
            'default_out': '大谷高校_新3年_購入票.xlsx'},
    },
    '清水丘高校': {
        '教科書マスターの照合': {
            'inputs': ['供給所データ（需要数データ）', '新2年の購入票Excel',
                       '新3年の購入票Excel', '新1年の購入票Excel'],
            'has_print': False, 'run': run_mdiff_shimizuoka, 'grade': 'マスター照合',
            'default_out': 'マスター照合_清水丘高校.xlsx'},

        '発注部数の照合（全学年）': {
            'inputs': ['供給所データ（需要数データ）', '新2年の購入票Excel',
                       '新3年の購入票Excel', '新1年の購入票Excel'],
            'numbers': [('1年生の人数', '全生徒')],
            'has_print': False, 'run': run_shortage_shimizuoka, 'grade': '照合',
            'default_out': '照合_清水丘高校_発注部数.xlsx'},

        '新1年（共通リスト）': {
            'inputs': ['新1年の購入票Excelファイル'],
            'has_print': True, 'run': run_shimizuoka_1nen, 'grade': '1',
            'default_out': '清水丘高校_新1年_購入リスト.xlsx'},
        '新2年': {
            'inputs': ['新2年の購入票Excelファイル'],
            'has_print': True, 'run': run_shimizuoka_students, 'grade': '2',
            'default_out': '清水丘高校_新2年_購入票.xlsx'},
        '新3年': {
            'inputs': ['新3年の購入票Excelファイル'],
            'has_print': True, 'run': run_shimizuoka_students, 'grade': '3',
            'default_out': '清水丘高校_新3年_購入票.xlsx'},
    },
    '工業高校': {
        '教科書マスターの照合': {
            'inputs': ['供給所データ（需要数データ）', '工業高校Excelファイル'],
            'has_print': False, 'run': run_mdiff_kogyo, 'grade': 'マスター照合',
            'default_out': 'マスター照合_工業高校.xlsx'},

        '発注部数の照合（全学年）': {
            'inputs': ['供給所データ（需要数データ）', '工業高校Excelファイル'],
            'numbers': [('1年 電気科', '1年電気科'), ('1年 建設科', '1年建設科'),
                        ('1年 電子機械科', '1年電子機械科'),
                        ('2年 電気科', '2年電気科'), ('2年 建設科（土木）', '2年建設科（土木）'),
                        ('2年 建設科（建築）', '2年建設科（建築）'), ('2年 電子機械科', '2年電子機械科'),
                        ('3年 環境土木科', '3年環境土木科'), ('3年 電気科（進学）', '3年電気科（進学）'),
                        ('3年 電気科（専門）', '3年電気科（専門）'), ('3年 建築科', '3年建築科'),
                        ('3年 電子機械科（進学）', '3年電子機械科（進学）'),
                        ('3年 電子機械科（専門）', '3年電子機械科（専門）')],
            'has_print': False, 'run': run_shortage_kogyo, 'grade': '照合',
            'default_out': '照合_工業高校_発注部数.xlsx'},

        '新1年（学科別リスト）': {
            'inputs': ['工業高校Excelファイル'],
            'has_print': True, 'run': run_kogyo, 'grade': '1',
            'default_out': '工業高校_新1年_学科別購入リスト.xlsx'},
        '新2年（学科別リスト）': {
            'inputs': ['工業高校Excelファイル'],
            'has_print': True, 'run': run_kogyo, 'grade': '2',
            'default_out': '工業高校_新2年_学科別購入リスト.xlsx'},
        '新3年（学科別リスト）': {
            'inputs': ['工業高校Excelファイル'],
            'has_print': True, 'run': run_kogyo, 'grade': '3',
            'default_out': '工業高校_新3年_学科別購入リスト.xlsx'},
    },
}

EXCEL_TYPES = [('Excelファイル', '*.xlsx *.xlsm'), ('すべてのファイル', '*.*')]


# ============================================================
# GUI本体
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('高校教科書購入票 自動生成ツール')
        self.geometry('820x720')
        self.input_vars = []
        self.number_vars = []
        self._build()

    # ---------- 画面の組み立て ----------
    def _build(self):
        pad = {'padx': 8, 'pady': 4}

        top = ttk.LabelFrame(self, text='① 学校と学年を選ぶ')
        top.pack(fill='x', **pad)

        ttk.Label(top, text='学校').grid(row=0, column=0, sticky='w', padx=8, pady=6)
        self.school_var = tk.StringVar(value=list(SCHOOLS.keys())[0])
        self.school_cb = ttk.Combobox(top, textvariable=self.school_var, state='readonly',
                                      values=list(SCHOOLS.keys()), width=28)
        self.school_cb.grid(row=0, column=1, sticky='w', padx=8, pady=6)
        self.school_cb.bind('<<ComboboxSelected>>', self._on_school_change)

        ttk.Label(top, text='学年').grid(row=0, column=2, sticky='w', padx=8, pady=6)
        self.grade_var = tk.StringVar()
        self.grade_cb = ttk.Combobox(top, textvariable=self.grade_var, state='readonly', width=30)
        self.grade_cb.grid(row=0, column=3, sticky='w', padx=8, pady=6)
        self.grade_cb.bind('<<ComboboxSelected>>', self._on_grade_change)

        self.in_frame = ttk.LabelFrame(self, text='② 入力ファイルを選ぶ')
        self.in_frame.pack(fill='x', **pad)

        self.num_frame = ttk.LabelFrame(self, text='②-2 名簿が無い学年の人数を入れる')
        self.num_frame.pack(fill='x', **pad)

        out_frame = ttk.LabelFrame(self, text='③ 出力先を決める')
        out_frame.pack(fill='x', **pad)
        ttk.Label(out_frame, text='出力ファイル').grid(row=0, column=0, sticky='w', padx=8, pady=6)
        self.out_var = tk.StringVar()
        ttk.Entry(out_frame, textvariable=self.out_var, width=68).grid(row=0, column=1, padx=4, pady=6)
        ttk.Button(out_frame, text='保存先...', command=self._pick_out).grid(row=0, column=2, padx=6)

        self.print_frame = ttk.LabelFrame(self, text='④ 印刷用レイアウト（1生徒1ページ）')
        self.print_frame.pack(fill='x', **pad)
        self.print_on = tk.BooleanVar(value=False)
        self.print_chk = ttk.Checkbutton(self.print_frame, text='印刷用レイアウトも作成する',
                                         variable=self.print_on, command=self._toggle_print)
        self.print_chk.grid(row=0, column=0, columnspan=3, sticky='w', padx=8, pady=4)

        ttk.Label(self.print_frame, text='印刷用ファイル').grid(row=1, column=0, sticky='w', padx=8, pady=4)
        self.print_out_var = tk.StringVar()
        self.print_out_entry = ttk.Entry(self.print_frame, textvariable=self.print_out_var, width=62)
        self.print_out_entry.grid(row=1, column=1, padx=4, pady=4)
        self.print_out_btn = ttk.Button(self.print_frame, text='保存先...', command=self._pick_print_out)
        self.print_out_btn.grid(row=1, column=2, padx=6)

        ttk.Label(self.print_frame, text='販売日').grid(row=2, column=0, sticky='w', padx=8, pady=4)
        self.sale_var = tk.StringVar()
        self.sale_entry = ttk.Entry(self.print_frame, textvariable=self.sale_var, width=62)
        self.sale_entry.grid(row=2, column=1, padx=4, pady=4)

        ttk.Label(self.print_frame, text='注意事項\n（1行に1つ）').grid(row=3, column=0, sticky='nw', padx=8, pady=4)
        self.notes_text = tk.Text(self.print_frame, height=4, width=62)
        self.notes_text.grid(row=3, column=1, padx=4, pady=4)

        btn_row = ttk.Frame(self)
        btn_row.pack(fill='x', **pad)
        self.run_btn = ttk.Button(btn_row, text='実行', command=self._run)
        self.run_btn.pack(side='left', padx=8)
        ttk.Button(btn_row, text='ログを消す', command=self._clear_log).pack(side='left')

        log_frame = ttk.LabelFrame(self, text='実行結果')
        log_frame.pack(fill='both', expand=True, **pad)
        self.log = tk.Text(log_frame, height=12, state='disabled')
        self.log.pack(side='left', fill='both', expand=True, padx=4, pady=4)
        sb = ttk.Scrollbar(log_frame, command=self.log.yview)
        sb.pack(side='right', fill='y')
        self.log.configure(yscrollcommand=sb.set)

        self._on_school_change()

    # ---------- 選択に応じた切り替え ----------
    def _on_school_change(self, event=None):
        grades = list(SCHOOLS[self.school_var.get()].keys())
        self.grade_cb['values'] = grades
        self.grade_var.set(grades[0])
        self._on_grade_change()

    def _on_grade_change(self, event=None):
        cfg = self._cfg()
        for child in self.in_frame.winfo_children():
            child.destroy()
        self.input_vars = []
        for i, label in enumerate(cfg['inputs']):
            ttk.Label(self.in_frame, text=label).grid(row=i, column=0, sticky='w', padx=8, pady=5)
            var = tk.StringVar()
            ttk.Entry(self.in_frame, textvariable=var, width=58).grid(row=i, column=1, padx=4, pady=5)
            ttk.Button(self.in_frame, text='参照...',
                       command=lambda v=var: self._pick_input(v)).grid(row=i, column=2, padx=6)
            self.input_vars.append(var)

        # 人数入力欄（照合メニューなど、名簿が無い学年がある場合だけ表示）
        for child in self.num_frame.winfo_children():
            child.destroy()
        self.number_vars = []
        nums = cfg.get('numbers', [])
        if nums:
            self.num_frame.pack(fill='x', padx=8, pady=4,
                                after=self.in_frame)
            ttk.Label(self.num_frame,
                      text='空欄のままでも実行できます（その学年は必要数に含めません）').grid(
                row=0, column=0, columnspan=6, sticky='w', padx=8, pady=(4, 2))
            for i, (label, key) in enumerate(nums):
                rr, cc = 1 + i // 3, (i % 3) * 2
                ttk.Label(self.num_frame, text=label).grid(row=rr, column=cc, sticky='e', padx=6, pady=3)
                var = tk.StringVar()
                ttk.Entry(self.num_frame, textvariable=var, width=8).grid(
                    row=rr, column=cc + 1, sticky='w', padx=(0, 12), pady=3)
                self.number_vars.append((key, var))
        else:
            self.num_frame.pack_forget()

        self.out_var.set(cfg.get('default_out', ''))
        base = os.path.splitext(cfg.get('default_out', 'output.xlsx'))[0]
        self.print_out_var.set(f'{base}（印刷用）.xlsx')

        if cfg['has_print']:
            self.print_chk.state(['!disabled'])
        else:
            self.print_on.set(False)
            self.print_chk.state(['disabled'])
        self._toggle_print()

    def _toggle_print(self):
        cfg = self._cfg()
        on = self.print_on.get() and cfg['has_print']
        state = 'normal' if on else 'disabled'
        self.print_out_entry.configure(state=state)
        self.print_out_btn.configure(state=state)
        self.sale_entry.configure(state=state)
        self.notes_text.configure(state=('normal' if on else 'disabled'))

    def _cfg(self):
        return SCHOOLS[self.school_var.get()][self.grade_var.get()]

    # ---------- ファイル選択 ----------
    def _pick_input(self, var):
        path = filedialog.askopenfilename(title='入力ファイルを選んでください', filetypes=EXCEL_TYPES)
        if path:
            var.set(path)

    def _pick_out(self):
        path = filedialog.asksaveasfilename(title='出力先', defaultextension='.xlsx',
                                            initialfile=self.out_var.get(),
                                            filetypes=[('Excelファイル', '*.xlsx')])
        if path:
            self.out_var.set(path)

    def _pick_print_out(self):
        path = filedialog.asksaveasfilename(title='印刷用ファイルの出力先', defaultextension='.xlsx',
                                            initialfile=self.print_out_var.get(),
                                            filetypes=[('Excelファイル', '*.xlsx')])
        if path:
            self.print_out_var.set(path)

    # ---------- ログ ----------
    def _write_log(self, text):
        self.log.configure(state='normal')
        self.log.insert('end', text)
        self.log.see('end')
        self.log.configure(state='disabled')

    def _clear_log(self):
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')

    # ---------- 実行 ----------
    def _run(self):
        cfg = self._cfg()
        paths = [v.get().strip() for v in self.input_vars]
        out_path = self.out_var.get().strip()

        for label, p in zip(cfg['inputs'], paths):
            if not p:
                messagebox.showwarning('入力不足', f'「{label}」が選ばれていません。')
                return
            if not os.path.exists(p):
                messagebox.showerror('ファイルなし', f'ファイルが見つかりません:\n{p}')
                return
        if not out_path:
            messagebox.showwarning('入力不足', '出力ファイル名を入れてください。')
            return

        popts = {'enabled': False, 'path': '', 'sale_date': '', 'notes': []}
        if self.print_on.get() and cfg['has_print']:
            print_out = self.print_out_var.get().strip()
            if not print_out:
                messagebox.showwarning('入力不足', '印刷用ファイル名を入れてください。')
                return
            notes = [ln.strip() for ln in self.notes_text.get('1.0', 'end').splitlines() if ln.strip()]
            popts = {'enabled': True, 'path': print_out,
                     'sale_date': self.sale_var.get().strip(), 'notes': notes}

        numbers = {}
        for key, var in self.number_vars:
            t = var.get().strip()
            if not t:
                continue
            if not t.isdigit():
                messagebox.showwarning('入力エラー', f'人数は数字で入れてください（{key}）。')
                return
            numbers[key] = int(t)

        self.run_btn.configure(state='disabled')
        self._write_log(f"\n=== {self.school_var.get()} / {self.grade_var.get()} 実行開始 ===\n")
        threading.Thread(target=self._worker, args=(cfg, paths, out_path, popts, numbers), daemon=True).start()

    def _worker(self, cfg, paths, out_path, popts, numbers=None):
        buf = io.StringIO()
        try:
            with redirect_stdout(buf):
                cfg['run'](paths, out_path, popts, cfg['grade'], numbers)
            msg = buf.getvalue() + '完了しました。\n'
            self.after(0, lambda: self._finish(msg, True))
        except Exception:
            msg = buf.getvalue() + '\n--- エラー ---\n' + traceback.format_exc()
            self.after(0, lambda: self._finish(msg, False))

    def _finish(self, msg, ok):
        self._write_log(msg)
        self.run_btn.configure(state='normal')
        if ok:
            messagebox.showinfo('完了', 'ファイルを作成しました。')
        else:
            messagebox.showerror('エラー', '処理中にエラーが発生しました。\n実行結果の欄をご確認ください。')


if __name__ == '__main__':
    App().mainloop()
