import pdfplumber
import fitz  # PyMuPDF
import pathlib
import gc
from enum import Enum

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
import datetime

import tkinter as tk
import tkinter.filedialog as filedialog
import requests

def check_typo_with_direct_url(sentence):
    url = "https://api.a3rt.recruit.co.jp/proofreading/v2/typo"
    params = {
        "apikey": "",
        "sentence": sentence,
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        content_type = response.headers.get('Content-Type')
        if 'application/json' in content_type:
            try:
                return response.json()
            except requests.exceptions.JSONDecodeError:
                print("レスポンスはJSON形式で解析できません。")
                return None
        else:
            print("レスポンスはJSON形式ではありません。")
            print("レスポンスの内容:")
            print(response.text)  # デバッグ用にレスポンスの内容を印刷
            return None
    else:
        print(f"HTTPエラー: {response.status_code}")
        return None


def output_report(doc, all_positions, pdf_path, log):

    total_pages = len(doc)  # PDFの総ページ数を取得
    if all_positions is not None:
        print(f"Found {len(all_positions)} issues.")
        for pos in all_positions:
            if pos['page'] < 0 or pos['page'] >= total_pages:
                print(f"Warning: Page {pos['page']} is out of range.")
                continue  # 範囲外のページ番号はスキップ
            print(f"Page: {pos['page']}, Character: {pos['character']}, Rect: {pos['rect']}")
            page = doc.load_page(pos['page'])
            rect = fitz.Rect(pos['rect'])
            highlight = page.add_highlight_annot(rect)
            highlight.update()

    new_filename = pathlib.Path(pdf_path).stem + "_highlighted.pdf"
    doc.save(new_filename)
    doc.close()
    # return log

def output_log(log):
    # PDFドキュメントを作成
    date = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    filename = f"sammary_report.pdf"
    doc = SimpleDocTemplate(filename, pagesize=letter)
    styles = getSampleStyleSheet()

    # PDFに追加する要素リスト
    elements = []

    # タイトルを追加
    title = f"Log Report - {date}"
    elements.append(Paragraph(title, styles['Title']))
    elements.append(Spacer(1, 12))

    # ログメッセージを追加
    for message in log:
        elements.append(Paragraph(message, styles['BodyText']))
        elements.append(Spacer(1, 12))

    # PDFを生成
    doc.build(elements)

class MARK_TYPES(Enum):
    JA_MARKS = 0    # "、","。"であることをチェックする場合 "," "."がハイライトされる
    EN_MARKS = 1    # ",","."であることをチェックする場合 "、" "。"がハイライトされる


def find_punctuation_positions(characters:str, log, doc):
    """
    charactersのある位置のRectのリストを返す
    """
    positions = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text_instances = page.search_for(characters)
        for inst in text_instances:
            positions.append({
                "page": page_num,
                "character": characters,
                "rect": inst
            })
            logtext = f"Page {positions['page']}: The punctuation mark '{positions['character']}' is incorrect."
            log.append(logtext)


    return positions, log

def find_all_punctuation_positions(check_marks_type:MARK_TYPES, log, doc):
    """
    すべてのページで",","."と"、","。"のチェックを行う
    """
    if MARK_TYPES.JA_MARKS == check_marks_type:
        highlight_punctuation_marks = [",", "."]
    elif MARK_TYPES.EN_MARKS == check_marks_type:
        highlight_punctuation_marks = ["、", "。"]
    else:
        raise ValueError("Invalid MARK_TYPES")

    all_positions = []

    for mark in highlight_punctuation_marks:
        positions, log = find_punctuation_positions(mark, log, doc)
        all_positions.extend(positions)

    return all_positions, log

def check_typo(doc, log):
    """
    ページごとに文字列を取得し、typoのチェックを行う
    """
    typo = []
    for page in range(len(doc)):
        page = doc.load_page(page)
        text = page.get_text()
        # 誤字チェック（API）何文字目に誤字があるかを返す
        result = check_typo_with_direct_url(text)
        alerts = result.get('alerts')

        if alerts is not None:
            for alert in alerts:
                typopoint = alert.get('pos')
                characters = alert.get('word')

                # その文字がpdfのどこの位置にあるかを求める
                rect = page.search_for(characters)

                typo.append({
                        "page": page,
                        "character": characters,
                        "rect": rect
                    })
        log.append(f"Page {page}: The character '{characters}' is incorrect.")

    return typo, log

def main(f_path:str,check_marks_type:MARK_TYPES):
    log = []
    print(f"Processing File: {f_path}")
    doc = fitz.open(f_path)
    all_positions, log = find_all_punctuation_positions(check_marks_type, log, doc)
    typo, log = check_typo(doc, log)
    output_report(doc, all_positions, f_path, log)
    output_log(log)



if __name__ == "__main__":
    f_path = "test.pdf"
    tk.Tk().withdraw()
    f_path = filedialog.askopenfilename()
    check_marks_type = MARK_TYPES.JA_MARKS #チェックしたいマークの種類を指定

    main(f_path,check_marks_type)
