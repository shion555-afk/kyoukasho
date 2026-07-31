# -*- coding: utf-8 -*-
"""
高校教科書購入票 自動生成ツール（全校共通メニュー）

使い方：
  python main.py
  もしくは実行ファイル化した場合はダブルクリックで起動してください。
"""
import sys

SCHOOLS = {
    '1': ('室蘭栄高校', 'saka'),
    '2': ('東翔高校', 'touhou'),
    '3': ('海星学院高校', 'kaisei'),
    '4': ('大谷高校（全日制）', 'otani'),
    '5': ('清水丘高校', 'shimizuoka'),
    '6': ('工業高校', 'kogyo'),
}


def main():
    print("========================================")
    print(" 高校教科書購入票 自動生成ツール")
    print("========================================")
    for key, (name, _) in SCHOOLS.items():
        print(f"{key}: {name}")
    choice = input("学校を選んでください: ").strip()

    if choice not in SCHOOLS:
        print("番号が正しくありません。1〜6を入力してください。")
        return

    name, module_name = SCHOOLS[choice]
    print(f"\n--- {name} ---\n")
    module = __import__(module_name)
    module.main()


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f"\nエラーが発生しました: {e}")
    input("\nEnterキーを押すと終了します...")
