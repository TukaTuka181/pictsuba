from pypdf import PdfReader
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.table import Table as DocxTable
import json
import re
import base64

def normalize(text: str) -> str:
    return re.sub(r'\s+', '', text).strip()


# --- 画像処理の関数群 ---

def has_inline_image(para) -> bool:
    """段落内に行内画像が含まれるか判定"""
    if para._element.findall('.//' + qn('w:drawing')):
      return True
    return False

def paragraph_to_dict(para, page_no, image_no, doc) -> dict:
    """段落を構造化して辞書で返す"""
    data = {
        "type": "image",
        "page_no": page_no,
        "image_no": image_no,
        "data_url": None,
        "paragraph_style": para.style.name if para.style else None,
        "alignment": para.alignment
    }

    for drawing in para._element.findall('.//' + qn('w:drawing')):
        blip = drawing.find('.//' + qn('a:blip'))
        if blip is None:
            continue
        r_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        if not r_id:
            continue
        image_part = doc.part.related_parts.get(r_id)
        if image_part is None:
            continue
        base64_str = base64.b64encode(image_part.blob).decode("utf-8")
        data["data_url"] = base64_str

    return data

def extract_images_from_para(para, doc) -> list:
    """段落内の画像をbase64に変換して返す"""
    images = []

    for drawing in para._element.findall('.//' + qn('w:drawing')):

        # rIdの取得（画像ファイルへの参照ID）
        blip = drawing.find('.//' + qn('a:blip'))
        if blip is None:
            continue
        r_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        if not r_id:
            continue

        # rIdから画像のバイナリデータを取得
        image_part = doc.part.related_parts.get(r_id)
        if image_part is None:
            continue

        # base64に変換
        image_data   = image_part.blob
        base64_str   = base64.b64encode(image_data).decode("utf-8")

    return base64_str

def find_page_no(text: str) -> int | None:
    """
    正規化テキストをPDFの残り文字列の先頭から探し、
    マッチしたらその部分を削除（消費）してページ番号を返す
    """
    norm = normalize(text)
    if not norm:
        return None

    for pdf_page in pdf_pages:
        idx = pdf_page["remaining"].find(norm)
        if idx != -1:
            # マッチした箇所より前の文字も消費（読み飛ばし）
            pdf_page["remaining"] = pdf_page["remaining"][idx + len(norm):]

            if not pdf_page["remaining"]:
              pdf_page["last_word"] = norm

            return pdf_page["page_no"]

    return None  # どのページにもマッチしなかった

# //////////////////////////////////////////////////////////////////////////
# //////////////////////////////////////////////////////////////////////////

# --- PDFからページごとの正規化テキストとポインタを準備 ---
pdf_reader = PdfReader("A製品契約書.pdf")
pdf_pages = []
for i, page in enumerate(pdf_reader.pages):
    text = normalize(page.extract_text() or "")

    pdf_pages.append({
        "page_no": i + 1,
        "remaining": text,
        "last_word": None
    })

# --- docxのブロック要素を出現順に処理 ---
doc = Document("A製品契約書.docx")
results = []
paragraph_no = 0
table_no = 0
image_no = 0

for child in doc.element.body.iterchildren():
    tag = child.tag

    # 段落
    if child.tag == qn('w:p'):
        para = DocxParagraph(child, doc)

        pPr = para._element.pPr
        if pPr is not None and pPr.sectPr is not None:
            results.append({
                "type": "section_break",
                "paragraph_no": paragraph_no,
                "text": ""
            })
            paragraph_no += 1
            continue


        # 画像を含む段落の処理
        if has_inline_image(para):

            # 前のresultからページ番号を推定
            find_result = next(
                (r for r in reversed(results) if r.get("page_no")),
                None
            )


            if not find_result:
                prev_page_no = None
            else:
                prev_page_no = find_result["page_no"]

                if find_result and "".join(find_result.get("text", "")) == pdf_pages[prev_page_no - 1]["last_word"]:
                    prev_page_no = None

            paragraph_data = paragraph_to_dict(para,prev_page_no, image_no, doc)

            results.append(paragraph_data)
            image_no += 1

        # runごとにテキストを取得して結合 & ページ番号を探す
        full_text = para.text.strip()

        
        if not full_text:
            continue

        # runを結合しながらページ番号を特定
        run_text = ""
        page_no = None
        for run in para.runs:
            run_text += run.text
            candidate = normalize(run_text)
            if candidate:
                result = find_page_no(run.text)
                if result and page_no is None:
                    page_no = result

        results.append({
            "type": "paragraph",
            "page_no": page_no,
            "paragraph_no": paragraph_no,
            "text": full_text
        })
        paragraph_no += 1

    # テーブル
    elif tag == qn('w:tbl'):
        table = DocxTable(child, doc)
        table_page_no = None

        table_data = []
        for row_no, row in enumerate(table.rows):
            row_data = []
            for col_no, cell in enumerate(row.cells):
                cell_text = cell.text.strip()
                row_data.append(cell_text)

            result = find_page_no(" ".join(row_data))
            if result:
                table_page_no = result
            table_data.append(row_data)

            results.append({
                "type": "table",
                "page_no": table_page_no,
                "table_no": table_no,
                "row": row_no,
                "text": row_data
            })
        table_no += 1

for i, item in enumerate(results):
    if item["type"] == "image" and item["page_no"] is None:
        # 文章の最後に画像がある場合、前のページ番号を引き継ぐ
        if not results[i+1:]:
            find_result = next(
                (r for r in reversed(results) if r.get("page_no")),
                None
            )
            prev_page_no = find_result["page_no"]
            item["page_no"] = prev_page_no
            continue    
        
        prev_page_no = next(
            (r["page_no"] for r in results[i+1:] if r.get("page_no")),
            None
        )
        item["page_no"] = prev_page_no

print(json.dumps(results, ensure_ascii=False, indent=2))
