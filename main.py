import fitz
import json
import argparse
import unicodedata
from rich.progress import Progress

# コマンドライン引数を解析
parser = argparse.ArgumentParser(description="PDFを再OCR処理するスクリプト")
parser.add_argument("--input_pdf_path", help="処理対象のPDFファイルのパス")
parser.add_argument("--output_pdf_path", help="処理後のPDFファイルのパス")
parser.add_argument(
    "--json_path", help="YomiTokuのOCR結果(JSON)のパス(`yomitoku {filename} -o {output_directory} -f json --combine`)")
args = parser.parse_args()

# JSONファイルを読み込み（ページごとの情報が配列になっている前提）
with open(args.json_path, "r", encoding="utf-8") as f:
    page_data = json.load(f)

# 元のPDFを読み込み、編集して保存する
input_pdf_path = args.input_pdf_path
output_pdf_path = args.output_pdf_path

doc = fitz.open(input_pdf_path)

# 既存テキストデータを全削除
print("既存のOCR結果を削除中...")
with Progress() as progress:
    task = progress.add_task("Processing spans", total=sum(
        1 for page_num in range(len(doc))
        for block in doc[page_num].get_textpage().extractDICT()["blocks"]
        for line in block["lines"]
        for span in line["spans"]
    ))
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)

        text_page = page.get_textpage()

        # ページからテキストとその位置を抽出
        blocks = text_page.extractDICT()["blocks"]

        # 各テキストブロックを処理
        for block in blocks:
            for line in block["lines"]:
                for span in line["spans"]:
                    # print(f"[{page_num+1} ページ目で検出] text:{span['text']}")
                    page.add_redact_annot(span["bbox"])
                    progress.update(task, advance=1)

        # https://stackoverflow.com/questions/72033672/delete-text-from-pdf-using-fitz
        page.apply_redactions(images=0)

# ページ数チェック
print("ページ数チェック中...")
if len(doc) != len(page_data):
    raise ValueError("PDFのページ数とJSONのデータ数が一致しません")

# テキストデータ書き込み処理
print("検索用テキストを処理中...")
with Progress() as progress:
    task = progress.add_task("Inserting textboxes", total=len(doc))
    for page_index, page in enumerate(doc):

        # JSONから追加するテキスト情報を取得
        words = page_data[page_index].get("words", [])

        for word in words:
            text = unicodedata.normalize('NFKC', word["content"])
            points = word["points"]
            direction = word["direction"]

            # px → pt に変換（DPIを考慮）
            def px_to_pt(x_px, y_px):
                dpi = 200
                x_pt = x_px * 72 / dpi
                y_pt = y_px * 72 / dpi
                return x_pt, y_pt

            # 座標変換（pointsは左上, 右上, 右下, 左下）
            x0_pt, y0_pt = px_to_pt(points[0][0], points[0][1])
            x1_pt, y1_pt = px_to_pt(points[1][0], points[1][1])
            x2_pt, y2_pt = px_to_pt(points[2][0], points[2][1])
            x3_pt, y3_pt = px_to_pt(points[3][0], points[3][1])

            x0_pt = min(x3_pt, x0_pt)
            x1_pt = max(x2_pt, x1_pt)
            y0_pt = min(y0_pt, y1_pt)
            y1_pt = max(y2_pt, y3_pt)

            left = x0_pt
            bottom = y1_pt
            right = x1_pt+9 if text.endswith("。") else x1_pt
            top = y0_pt

            rect = fitz.Rect(left, top, right, bottom)

            # 自動調整フォントサイズの計算
            base_fontsize = 10  # 適当な初期サイズ
            text_width_at_base = fitz.get_text_length(
                text, fontname="japan-s", fontsize=base_fontsize)
            if text_width_at_base == 0:
                fontsize = 5  # 適当な最小サイズを設定（空文字などの保険）
            else:
                scale = (rect.height if direction ==
                         "vertical"else rect.width) / text_width_at_base
                fontsize = base_fontsize * scale

            # 1だとギリギリrect外に出て描画されないことがあるので小さくする
            fontsize = fontsize * 0.99

            # テキストの表示開始Y座標（下寄せ）
            y_start = rect.y1 - fontsize-3
            rect = fitz.Rect(rect.x0, y_start, rect.x1, rect.y1)

            # print(f"[ページ{page_index+1}] テキスト: {text}")

            # page.draw_rect(rect, color=(1, 0, 0), width=0.5)

            # 透明テキストを描画
            page.insert_textbox(
                rect, text,
                fontsize=fontsize,
                color=(0.5, 0.5, 0.5),
                render_mode=3,
                overlay=True,
                fontname="japan-s",
                rotate=270 if direction == "vertical" else 0,
            )

        progress.update(task, advance=1)

# 保存
print("PDFを保存中...")
doc.save(output_pdf_path, garbage=4, deflate=True)
doc.close()

print("正常に完了しました！")
