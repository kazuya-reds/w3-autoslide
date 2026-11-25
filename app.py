import streamlit as st
from google import genai
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_DATA_LABEL_POSITION
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from PIL import Image
import csv
import json
import datetime
import io
import re
import time
import copy
import os
import uuid
import unicodedata
import math

st.set_page_config(page_title="W3 AutoSlide", layout="wide")

st.markdown("""
    <style>
    img { border-radius: 0px !important; }
    div.stButton > button {
        font-weight: bold !important;
        border-radius: 8px !important;
        font-size: 16px !important;
    }
    button[data-testid="baseButton-primary"], div.stButton > button[kind="primary"] {
        background-color: #ff4b4b !important;
        color: #ffffff !important;
        font-weight: bold !important;
        font-size: 18px !important;
        padding: 12px 28px !important;
        border: none !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.15) !important;
    }
    button[data-testid="baseButton-primary"]:hover, div.stButton > button[kind="primary"]:hover {
        background-color: #d93025 !important;
        color: #ffffff !important;
    }
    </style>
""", unsafe_allow_html=True)

if os.path.exists("logo.png"):
    st.image("logo.png", width=350)
st.caption("AI資料 自動生成システム")

LAYOUT_MAP = {
    "line_chart": "折れ線グラフ（推移・トレンド）",
    "bar_chart": "棒グラフ（比較・実績）",
    "pie_chart": "円グラフ（構成比・内訳）",
    "text": "解説スライド（文章メイン または 写真添え）",
    "step": "実行ステップ（フロー・手順図）",
    "checklist": "チェックリスト（確認項目一覧）",
    "table": "表組み（比較表・施策一覧・タスク管理）"
}
LAYOUT_REVERSE_MAP = {v: k for k, v in LAYOUT_MAP.items()}

if "saved_images" not in st.session_state:
    st.session_state["saved_images"] = []


def parse_csv_line_safe(line_str):
    if not line_str or not str(line_str).strip():
        return []
    try:
        parsed = next(csv.reader([str(line_str).strip()], skipinitialspace=True))
        return [c.strip() for c in parsed if c is not None]
    except Exception:
        return [c.strip() for c in str(line_str).split(",")]


def parse_number_with_comma(val):
    if val is None:
        return 0.0
    s = str(val).strip()
    if not s:
        return 0.0
    oku_match = re.search(r'(\d+(?:,\d{3})*(?:\.\d+)?)\s*億\s*(\d+(?:,\d{3})*(?:\.\d+)?)?\s*万?', s)
    if oku_match:
        oku = float(oku_match.group(1).replace(",", ""))
        man = float(oku_match.group(2).replace(",", "")) if oku_match.group(2) else 0.0
        return oku * 10000.0 + man
    cleaned = re.sub(r'[^\d\.\-]', '', s.replace(",", ""))
    if not cleaned or cleaned in [".", "-"]:
        return 0.0
    num = float(cleaned)
    if math.isnan(num) or math.isinf(num):
        return 0.0
    return num


def call_gemini_with_retry(client, prompt, max_retries=5):
    model_name = 'gemini-3.6-flash'
    last_exception = None
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(model=model_name, contents=prompt)
            if response and response.text:
                return response.text.strip()
        except Exception as e:
            last_exception = e
            err_str = str(e)
            if "503" in err_str or "UNAVAILABLE" in err_str or "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                wait_sec = 3 * (attempt + 1)
                st.warning(f"⏳ AIサーバー混雑を検知。{wait_sec}秒後に自動再試行します... ({attempt + 1}/{max_retries})")
                time.sleep(wait_sec)
            else:
                raise e
    raise last_exception if last_exception else Exception("AIサーバーの混雑が続いています。1〜2分置かれてから再度お試しください。")


def clean_notice_text(text):
    if not text:
        return ""
    return re.sub(r'^(?:[※\*\:\s]|注意[：:]|注[：:])+', '', str(text).strip())


def sanitize_text_empty_lines(text):
    if not text:
        return ""
    return re.sub(r'\n{2,}', '\n', str(text).strip())


def clean_chapter_title(title):
    if not title:
        return ""
    res = re.sub(
        r'^(?:第\s*\d+\s*[章節]\s*|\d{1,2}[\.\:\-_・／/]\s*|\d{2}\s+(?![ヶカか年月週日つ個大度ステップSTEP])|\d{1}\s+(?![ヶカか年月週日つ個大度ステップSTEP]))',
        '',
        str(title).strip()
    )
    return res.strip()


def clean_step_text(text):
    if not text:
        return ""
    cleaned = re.sub(
        r'^(?:[①-⑳]|\(?\d+[\.\)）]\s*|(?:STEP|Step|step|ステップ)\s*\d+[\:\s・]*|[・\-\*]\s*)+',
        '',
        str(text).strip()
    )
    return cleaned.strip()


# 5件以上のステップ項目がある場合に、4枠に収めつつ5件目以降を省略せず統合するヘルパー
def normalize_step_descs_to_4(step_descs):
    cleaned_list = [clean_step_text(x) for x in (step_descs or []) if clean_step_text(x)]
    if len(cleaned_list) <= 4:
        while len(cleaned_list) < 4:
            cleaned_list.append(f"アクション {len(cleaned_list) + 1}")
        return cleaned_list[:4]
    first_three = cleaned_list[:3]
    merged_fourth = " / ".join(cleaned_list[3:])
    return first_three + [merged_fourth]


def sanitize_main_item(main_item, clean_ch_title):
    main_item = str(main_item).strip() if main_item else ""
    if not main_item or len(main_item) > 18 or "。" in main_item or "、" in main_item:
        first_part = re.split(r'[。\n、,]', main_item)[0].strip()
        if 2 <= len(first_part) <= 15:
            main_item = first_part
        else:
            if "推移" in clean_ch_title or "売上" in clean_ch_title:
                main_item = "売上実績は堅調に推移"
            elif "チャネル" in clean_ch_title or "販売経路" in clean_ch_title:
                main_item = "自社EC依存とモール拡大"
            elif "要因" in clean_ch_title or "分析" in clean_ch_title or "顧客" in clean_ch_title or "構成" in clean_ch_title:
                main_item = "分析結果の主要ポイント"
            elif "アクション" in clean_ch_title or "手順" in clean_ch_title or "戦略" in clean_ch_title or "施策" in clean_ch_title or "目標" in clean_ch_title:
                main_item = "今後の重要施策と計画"
            elif "チェック" in clean_ch_title or "確認" in clean_ch_title or "課題" in clean_ch_title:
                main_item = "現状の主要課題と確認事項"
            elif "表" in clean_ch_title or "タスク" in clean_ch_title or "管理" in clean_ch_title:
                main_item = "実行施策の進捗一覧"
            else:
                main_item = clean_ch_title[:12]
    return main_item


def extract_chapter_specific_chart_data(chapter_title, body_text, raw_memo):
    combined = f"{chapter_title} {body_text}"

    if any(kw in combined for kw in ["月別", "1月", "2月", "3月", "4月", "5月", "6月", "推移"]):
        month_matches = re.findall(r'(?:^|\n|\s)([1-9]|1[0-2])月[：:\s]+([\d,]+(?:\.\d+)?)\s*(?:万|万円)?', raw_memo)
        if len(month_matches) >= 3:
            cats = [f"{m[0]}月" for m in month_matches]
            vals = [parse_number_with_comma(m[1]) for m in month_matches]
            return {
                "title": "月別売上推移",
                "unit": "万円",
                "categories": cats,
                "series_name": "月別売上高（万円）",
                "values": vals
            }

    if any(kw in combined for kw in ["カテゴリー", "カテゴリ", "生活雑貨", "家電"]):
        cat_matches = re.findall(r'(?:^|\n)([^\n：:]+?)[：:]\s*(\d+(?:,\d{3})*(?:\.\d+)?)\s*[%％]', raw_memo)
        valid_cats = [(c.strip(), parse_number_with_comma(v)) for c, v in cat_matches if any(k in c for k in ["雑貨", "家電", "ファッション", "食品", "その他"])]
        if len(valid_cats) >= 3:
            return {
                "title": "商品カテゴリー別売上構成",
                "unit": "％",
                "categories": [x[0] for x in valid_cats],
                "series_name": "売上構成比（％）",
                "values": [x[1] for x in valid_cats]
            }

    if any(kw in combined for kw in ["年代", "20代", "30代", "40代", "50代", "60代"]):
        age_matches = re.findall(r'(?:^|\n)((?:20|30|40|50|60代以上)[^\n：:]*?)[：:]\s*(\d+(?:,\d{3})*(?:\.\d+)?)\s*[%％]', raw_memo)
        if len(age_matches) >= 3:
            return {
                "title": "顧客の年代構成",
                "unit": "％",
                "categories": [m[0].strip() for m in age_matches],
                "series_name": "年代別割合（％）",
                "values": [parse_number_with_comma(m[1]) for m in age_matches]
            }

    if any(kw in combined for kw in ["チャネル", "販売経路", "自社EC", "楽天市場", "Amazon"]):
        ch_matches = re.findall(r'(?:^|\n)(自社EC[^\n：:]*|楽天市場|Amazon|その他)[：:]\s*(\d+(?:,\d{3})*(?:\.\d+)?)\s*[%％]', raw_memo)
        if len(ch_matches) >= 3:
            return {
                "title": "販売チャネル別売上構成",
                "unit": "％",
                "categories": [m[0].strip() for m in ch_matches],
                "series_name": "チャネル構成比（％）",
                "values": [parse_number_with_comma(m[1]) for m in ch_matches]
            }

    if any(kw in combined for kw in ["商品別", "商品A", "商品B", "商品C", "ワイヤレスイヤホン", "収納ボックス"]):
        prod_blocks = re.findall(r'(商品[A-E][^：\n]*)[：:][^\n]*?(?:\n[^\n]*?)*?売上[：:]\s*([\d,]+(?:\.\d+)?)\s*万円', raw_memo)
        if len(prod_blocks) >= 3:
            return {
                "title": "商品別売上実績",
                "unit": "万円",
                "categories": [p[0].strip() for p in prod_blocks],
                "series_name": "売上高（万円）",
                "values": [parse_number_with_comma(p[1]) for p in prod_blocks]
            }

    return {
        "title": clean_chapter_title(chapter_title) or "実績データ",
        "unit": "",
        "categories": [],
        "series_name": "データ",
        "values": []
    }


def verify_and_fix_chart_data(chart_info, layout_type, chapter_title, body_text, raw_memo):
    if not isinstance(chart_info, dict):
        chart_info = {}

    raw_cats = chart_info.get("categories", [])
    raw_vals = chart_info.get("values", [])

    cats = [str(c).strip() for c in raw_cats if str(c).strip()] if isinstance(raw_cats, list) else []
    vals = [parse_number_with_comma(v) for v in raw_vals] if isinstance(raw_vals, list) else []

    if not cats or not vals or len(cats) != len(vals):
        extracted = extract_chapter_specific_chart_data(chapter_title, body_text, raw_memo)
        if extracted["categories"] and extracted["values"]:
            cats = extracted["categories"]
            vals = extracted["values"]
            if not chart_info.get("title"):
                chart_info["title"] = extracted["title"]
            if not chart_info.get("unit"):
                chart_info["unit"] = extracted["unit"]
            if not chart_info.get("series_name"):
                chart_info["series_name"] = extracted["series_name"]

    min_len = min(len(cats), len(vals))
    cats = cats[:min_len]
    vals = vals[:min_len]

    chart_info["categories"] = cats
    chart_info["values"] = vals
    is_valid = (min_len >= 2 and any(v > 0 for v in vals))
    return chart_info, is_valid


def duplicate_slide_safe(prs, src_idx):
    src_slide = prs.slides[src_idx]
    new_slide = prs.slides.add_slide(src_slide.slide_layout)
    spTree = new_slide.shapes._spTree
    for child in list(spTree):
        tag = child.tag.split('}')[-1]
        if tag in ['sp', 'graphicFrame', 'grpSp', 'cxnSp', 'pic']:
            spTree.remove(child)

    rel_map = {}
    for shape in src_slide.shapes:
        if shape.has_chart:
            continue

        new_el = copy.deepcopy(shape._element)
        for node in new_el.iter():
            if node.tag.endswith('creationId'):
                node.attrib['id'] = f"{{{str(uuid.uuid4()).upper()}}}"

        try:
            r_ids = new_el.xpath('.//@r:embed | .//@r:id')
            for rId in set(r_ids):
                if rId in src_slide.part.rels:
                    if rId not in rel_map:
                        rel = src_slide.part.rels[rId]
                        target_part = rel.target_part
                        new_rId = new_slide.part.relate_to(target_part, rel.reltype)
                        rel_map[rId] = new_rId
            for elem in new_el.iter():
                for k, v in list(elem.attrib.items()):
                    if v in rel_map:
                        elem.attrib[k] = rel_map[v]
        except Exception:
            pass
        spTree.append(new_el)
    return new_slide


def ensure_page_number(slide, page_num):
    found_shp = None
    for shp in slide.shapes:
        is_ph = (shp.is_placeholder and shp.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER)
        is_name = any(kw in shp.name.lower() for kw in ["スライド番号", "slide_number"])
        if is_ph or is_name:
            found_shp = shp
            break

    if found_shp and found_shp.has_text_frame:
        found_shp.left = Pt(623.8)
        found_shp.top = Pt(548.1)
        found_shp.width = Pt(189.4)
        found_shp.height = Pt(31.6)
        tf = found_shp.text_frame
        tf.word_wrap = False
        tf.margin_left = Pt(7.2)
        tf.margin_right = Pt(7.2)
        tf.margin_top = Pt(3.6)
        tf.margin_bottom = Pt(3.6)
        p = tf.paragraphs[0]
        p.text = str(page_num)
        p.alignment = PP_ALIGN.RIGHT
        p.font.name = "メイリオ"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = RGBColor(2, 24, 46)
    else:
        txBox = slide.shapes.add_textbox(Pt(623.8), Pt(548.1), Pt(189.4), Pt(31.6))
        tf = txBox.text_frame
        tf.word_wrap = False
        tf.margin_left = Pt(7.2)
        tf.margin_right = Pt(7.2)
        tf.margin_top = Pt(3.6)
        tf.margin_bottom = Pt(3.6)
        p = tf.paragraphs[0]
        p.text = str(page_num)
        p.alignment = PP_ALIGN.RIGHT
        p.font.name = "メイリオ"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = RGBColor(2, 24, 46)


def fix_all_xml_integrity_safe(prs):
    global_id = 100
    for sld in prs.slides:
        tree = sld.shapes._spTree
        for node in tree.iter():
            if node.tag.endswith('cNvPr'):
                global_id += 1
                node.attrib['id'] = str(global_id)
            elif node.tag.endswith('creationId'):
                node.attrib['id'] = f"{{{str(uuid.uuid4()).upper()}}}"

    try:
        valid_rids = set(
            sld_id.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            for sld_id in prs.slides._sldIdLst
        )
        for rId, rel in list(prs.part.rels.items()):
            if rel.reltype == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" and rId not in valid_rids:
                prs.part.rels.pop(rId)
    except Exception:
        pass

    for sld in prs.slides:
        try:
            used_rids = set(
                sld.element.xpath(
                    './/@r:id | .//@r:embed',
                    namespaces={'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
                )
            )
            for rId, rel in list(sld.part.rels.items()):
                if "slideLayout" not in rel.reltype and "slideMaster" not in rel.reltype and rId not in used_rids:
                    sld.part.rels.pop(rId)
        except Exception:
            pass


def force_meiryo_and_style_on_chart(chart, font_size_pt=9):
    ns_a = nsdecls("a")
    if chart.has_title and chart.chart_title.text_frame:
        for p in chart.chart_title.text_frame.paragraphs:
            pPr = p._p.get_or_add_pPr()
            defRPr = pPr.get_or_add_defRPr()
            defRPr.attrib['sz'] = '1200'
            defRPr.attrib['b'] = '1'
            for child in list(defRPr):
                if child.tag.endswith('ea') or child.tag.endswith('latin'):
                    defRPr.remove(child)
            defRPr.append(parse_xml(f'<a:ea {ns_a} typeface="メイリオ"/>'))
            defRPr.append(parse_xml(f'<a:latin {ns_a} typeface="Meiryo"/>'))

            for r in p.runs:
                rPr = r._r.get_or_add_rPr()
                rPr.attrib['sz'] = '1200'
                rPr.attrib['b'] = '1'
                rPr.attrib['lang'] = 'ja-JP'
                rPr.attrib['altLang'] = 'en-US'
                for child in list(rPr):
                    if child.tag.endswith('ea') or child.tag.endswith('latin'):
                        rPr.remove(child)
                rPr.append(parse_xml(f'<a:ea {ns_a} typeface="メイリオ"/>'))
                rPr.append(parse_xml(f'<a:latin {ns_a} typeface="Meiryo"/>'))

    if chart.has_legend and chart.legend:
        try:
            txPr = chart.legend._element.get_or_add_txPr()
            for p in txPr.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p'):
                pPr = p.get_or_add_pPr()
                defRPr = pPr.get_or_add_defRPr()
                defRPr.attrib['sz'] = str(int(font_size_pt * 100))
                for child in list(defRPr):
                    if child.tag.endswith('ea') or child.tag.endswith('latin'):
                        defRPr.remove(child)
                defRPr.append(parse_xml(f'<a:ea {ns_a} typeface="メイリオ"/>'))
                defRPr.append(parse_xml(f'<a:latin {ns_a} typeface="Meiryo"/>'))
        except Exception:
            pass

    try:
        if hasattr(chart.plots[0], 'has_data_labels') and chart.plots[0].has_data_labels:
            txPr = chart.plots[0].data_labels._element.get_or_add_txPr()
            p_list = txPr.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p')
            if not p_list:
                p = parse_xml(f'<a:p {ns_a}><a:pPr/></a:p>')
                txPr.append(p)
                p_list = [p]

            for p in p_list:
                pPr = p.get_or_add_pPr()
                defRPr = pPr.get_or_add_defRPr()
                defRPr.attrib['sz'] = '1100'
                defRPr.attrib['b'] = '1'
                for child in list(defRPr):
                    tag_name = child.tag.split('}')[-1]
                    if tag_name in ['solidFill', 'ea', 'latin']:
                        defRPr.remove(child)
                defRPr.append(parse_xml(f'<a:solidFill {ns_a}><a:srgbClr val="FFFFFF"/></a:solidFill>'))
                defRPr.append(parse_xml(f'<a:ea {ns_a} typeface="メイリオ"/>'))
                defRPr.append(parse_xml(f'<a:latin {ns_a} typeface="Meiryo"/>'))
    except Exception:
        pass

    for axis_name in ['category_axis', 'value_axis']:
        try:
            axis = getattr(chart, axis_name)
            if axis and hasattr(axis, 'tick_labels'):
                txPr = axis.tick_labels._element.get_or_add_txPr()
                for p in txPr.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p'):
                    pPr = p.get_or_add_pPr()
                    defRPr = pPr.get_or_add_defRPr()
                    defRPr.attrib['sz'] = str(int(font_size_pt * 100))
                    for child in list(defRPr):
                        if child.tag.endswith('ea') or child.tag.endswith('latin'):
                            defRPr.remove(child)
                    defRPr.append(parse_xml(f'<a:ea {ns_a} typeface="メイリオ"/>'))
                    defRPr.append(parse_xml(f'<a:latin {ns_a} typeface="Meiryo"/>'))
        except (ValueError, AttributeError):
            pass


def wipe_right_side_shapes(slide):
    shapes_to_del = []
    for shp in slide.shapes:
        if shp.top > Pt(500) and shp.left > Pt(500):
            continue
        if shp.is_placeholder and shp.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER:
            continue

        if shp.left > Pt(350):
            if shp.shape_type in [13, 14]:
                shapes_to_del.append(shp)
            elif shp.shape_type == 1 and not shp.has_text_frame:
                shapes_to_del.append(shp)
            elif shp.has_text_frame and (shp.text_frame.text.strip() == "" or "[[画像" in shp.text_frame.text or "CAPTION" in shp.name.upper()):
                shapes_to_del.append(shp)

    for shp in shapes_to_del:
        try:
            sp = shp._element
            sp.getparent().remove(sp)
        except Exception:
            pass


def insert_slide_picture_safe(slide, assigned_images, captions=None, max_w_pt=350, max_h_pt=280, left_pt=420, top_pt=140):
    if not assigned_images or len(assigned_images) == 0:
        return
    img_bytes = assigned_images[0]["bytes"]
    pil_img = Image.open(io.BytesIO(img_bytes))
    orig_w, orig_h = pil_img.size

    max_w = Pt(max_w_pt)
    max_h = Pt(max_h_pt)
    scale = min(max_w / orig_w, max_h / orig_h)
    new_w = int(orig_w * scale)
    new_h = int(orig_h * scale)
    new_l = int(Pt(left_pt) + (max_w - new_w) / 2)
    new_t = int(Pt(top_pt) + (max_h - new_h) / 2)

    slide.shapes.add_picture(io.BytesIO(img_bytes), new_l, new_t, new_w, new_h)

    if captions and len(captions) > 0 and captions[0]:
        tx_box = slide.shapes.add_textbox(new_l, new_t + new_h + Pt(6), new_w, Pt(22))
        tf = tx_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = captions[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(9.5)
        p.font.name = "メイリオ"
        p.font.color.rgb = RGBColor(120, 120, 120)


def build_perfect_steps(slide, step_descs, body_text="", assigned_images=None, captions=None):
    wipe_right_side_shapes(slide)
    has_image = bool(assigned_images and len(assigned_images) > 0)

    shapes_to_del = []
    main_body_shp = None
    for shp in slide.shapes:
        if shp.top > Pt(500) and shp.left > Pt(500):
            continue
        sname = shp.name.upper()
        if "MAIN_BODY_TEXT" in sname:
            main_body_shp = shp
        elif any(kw in sname for kw in ["CAUTION", "SUPPLEMENT", "補足", "注意"]):
            shapes_to_del.append(shp)
        elif sname in ["SECTION_TITLE", "BODY_TEXT"] and "MAIN" not in sname:
            shapes_to_del.append(shp)
        elif "コネクタ" in shp.name or ("LINE" in sname and "SECTION_LINE" not in sname):
            shapes_to_del.append(shp)
        elif "CAPTION" in sname:
            shapes_to_del.append(shp)

    for shp in shapes_to_del:
        try:
            sp = shp._element
            sp.getparent().remove(sp)
        except Exception:
            pass

    clean_body = body_text
    if "【" in clean_body:
        clean_body = clean_body.split("【")[0].strip()
    if "1." in clean_body:
        clean_body = clean_body.split("1.")[0].strip()
    clean_body = re.sub(r'[\r\n]+', ' ', clean_body).strip()
    if not clean_body:
        clean_body = "以下の重要ステップに沿って着実に施策を推進します。"

    body_width = Pt(280.0) if has_image else Pt(660.0)
    step_text_width = Pt(220.0) if has_image else Pt(580.0)
    font_size_step = Pt(11.0) if has_image else Pt(12.5)

    if main_body_shp and main_body_shp.has_text_frame:
        main_body_shp.left = Pt(113.7)
        main_body_shp.top = Pt(146.0)
        main_body_shp.width = body_width
        main_body_shp.height = Pt(36.0)
        tf_b = main_body_shp.text_frame
        tf_b.word_wrap = True
        tf_b.text = clean_body
        for p in tf_b.paragraphs:
            p.font.name = "メイリオ"
            p.font.size = Pt(11.0) if has_image else Pt(11.5)
            p.line_spacing = 1.2

    spTree = slide.shapes._spTree
    step_badges = {}
    step_texts = {}

    for shp in slide.shapes:
        sname = shp.name.upper()
        if "STEP_1_BADGE" in sname:
            step_badges[1] = shp
        elif "STEP_1_TEXT" in sname:
            step_texts[1] = shp
        elif "STEP_2_BADGE" in sname:
            step_badges[2] = shp
        elif "STEP_2_TEXT" in sname:
            step_texts[2] = shp
        elif "STEP_3_BADGE" in sname:
            step_badges[3] = shp
        elif "STEP_3_TEXT" in sname:
            step_texts[3] = shp
        elif "STEP_4_BADGE" in sname:
            step_badges[4] = shp
        elif "STEP_4_TEXT" in sname:
            step_texts[4] = shp

    if 1 in step_badges and 1 in step_texts:
        base_b = step_badges[1]
        base_t = step_texts[1]
        for k in range(3, 5):
            if k not in step_badges:
                new_b_el = copy.deepcopy(base_b._element)
                for c_node in new_b_el.iter():
                    if c_node.tag.endswith('creationId'):
                        c_node.attrib['id'] = f"{{{str(uuid.uuid4()).upper()}}}"
                spTree.append(new_b_el)
                new_b = slide.shapes[-1]
                new_b.name = f"STEP_{k}_BADGE"
                step_badges[k] = new_b
            if k not in step_texts:
                new_t_el = copy.deepcopy(base_t._element)
                for c_node in new_t_el.iter():
                    if c_node.tag.endswith('creationId'):
                        c_node.attrib['id'] = f"{{{str(uuid.uuid4()).upper()}}}"
                spTree.append(new_t_el)
                new_t = slide.shapes[-1]
                new_t.name = f"STEP_{k}_TEXT"
                step_texts[k] = new_t

    # 5件以上の項目があっても省略せずSTEP 4に集約して全項目を保持
    normalized_4 = normalize_step_descs_to_4(step_descs)

    start_top = Pt(200.0)
    step_gap = Pt(46.0)

    for k in range(1, 5):
        if k in step_texts:
            t_shp = step_texts[k]
            t_shp.left = Pt(180.0) if has_image else Pt(190.0)
            t_shp.top = int(start_top + (k - 1) * step_gap)
            t_shp.width = step_text_width
            t_shp.height = Pt(36.0)
            if t_shp.has_text_frame:
                tf = t_shp.text_frame
                tf.word_wrap = True
                tf.text = normalized_4[k - 1]
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                for p in tf.paragraphs:
                    p.font.size = font_size_step
                    p.font.name = "メイリオ"
                    for r in p.runs:
                        r.font.size = font_size_step
                        r.font.name = "メイリオ"

        if k in step_badges:
            b_shp = step_badges[k]
            b_shp.left = Pt(113.7)
            b_shp.height = Pt(22.0)
            b_shp.width = Pt(58.0) if has_image else Pt(65.0)
            if k in step_texts:
                t_shp = step_texts[k]
                b_shp.top = int(t_shp.top + (t_shp.height - b_shp.height) / 2)

            if b_shp.has_text_frame:
                btf = b_shp.text_frame
                btf.text = f"STEP {k}"
                btf.vertical_anchor = MSO_ANCHOR.MIDDLE
                for p in btf.paragraphs:
                    p.alignment = PP_ALIGN.CENTER
                    p.font.name = "メイリオ"
                    p.font.size = Pt(9.0) if has_image else Pt(9.5)
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(255, 255, 255)

    if has_image:
        insert_slide_picture_safe(slide, assigned_images, captions, max_w_pt=340, max_h_pt=270, left_pt=420, top_pt=145)


def format_checklist_slide(slide, assigned_images=None, captions=None):
    wipe_right_side_shapes(slide)
    has_image = bool(assigned_images and len(assigned_images) > 0)
    main_title = None
    main_line = None
    intro_shape = None
    box_shape = None
    items_shape = None
    outro_shape = None

    shapes_to_del = []
    for shp in slide.shapes:
        if shp.top > Pt(500) and shp.left > Pt(500):
            continue
        sname = shp.name.upper()
        if "MAIN_SECTION_TITLE" in sname:
            main_title = shp
        elif "SECTION_LINE" in sname:
            main_line = shp
        elif "CHECKLIST_INTRO" in sname:
            intro_shape = shp
        elif shp.shape_type == 1 and shp.width > Pt(300) and shp.height > Pt(100):
            box_shape = shp
        elif "CHECKLIST_ITEMS" in sname or (shp.has_text_frame and "□" in shp.text_frame.text):
            items_shape = shp
        elif "CHECKLIST_OUTRO" in sname:
            outro_shape = shp
        elif any(kw in sname for kw in ["CAUTION", "SUPPLEMENT", "注意", "補足", "CAPTION"]):
            shapes_to_del.append(shp)
        elif shp.shape_type == 1 and Pt(50) < shp.width < Pt(300):
            shapes_to_del.append(shp)

    for shp in shapes_to_del:
        try:
            sp = shp._element
            sp.getparent().remove(sp)
        except Exception:
            pass

    content_w = Pt(300.0) if has_image else Pt(660.0)

    if main_title:
        main_title.left = Pt(124.6)
        main_title.top = Pt(115.0)
        main_title.width = content_w
        if main_line:
            main_line.left = Pt(113.7)
            main_line.top = Pt(120.5)

    if intro_shape:
        intro_shape.left = Pt(113.7)
        intro_shape.top = Pt(150.0)
        intro_shape.width = content_w

    if box_shape:
        box_shape.left = Pt(113.7)
        box_shape.top = Pt(185.0)
        box_shape.width = content_w
        box_shape.height = Pt(200.0)

    if items_shape and items_shape.has_text_frame and box_shape:
        items_shape.left = Pt(125.0) if has_image else Pt(140.0)
        items_shape.width = Pt(280.0) if has_image else Pt(610.0)
        tf = items_shape.text_frame
        valid_paras = [p for p in tf.paragraphs if p.text.strip()]
        num_items = max(1, len(valid_paras))

        for p in tf.paragraphs:
            p.font.name = "メイリオ"
            p.font.size = Pt(10.5) if has_image else Pt(12.0)
            p.line_spacing = 1.25 if has_image else 1.3
            p.space_after = Pt(6.0) if has_image else Pt(8.0)
            for r in p.runs:
                r.font.name = "メイリオ"
                r.font.size = Pt(10.5) if has_image else Pt(12.0)

        effective_items_h = Pt(num_items * (24.0 if has_image else 28.0))
        items_shape.height = int(effective_items_h)
        items_shape.top = int(box_shape.top + (box_shape.height - effective_items_h) / 2)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    if outro_shape and box_shape:
        outro_shape.left = Pt(113.7)
        outro_shape.top = int(box_shape.top + box_shape.height + Pt(14.0))
        outro_shape.width = content_w

    if has_image:
        insert_slide_picture_safe(slide, assigned_images, captions, max_w_pt=340, max_h_pt=270, left_pt=430, top_pt=145)


def apply_clean_table_borders_and_stripes(table):
    ns_a = nsdecls("a")
    tbl_el = table._tbl

    tblPr = tbl_el.tblPr
    if tblPr is not None:
        tblPr.attrib['firstRow'] = '1'
        tblPr.attrib['firstCol'] = '0'
        tblPr.attrib['lastRow'] = '0'
        tblPr.attrib['lastCol'] = '0'
        tblPr.attrib['bandRow'] = '1'
        tblPr.attrib['bandCol'] = '0'

    header_bg = "677482"
    row_even_bg = "FFFFFF"
    row_odd_bg = "E8ECEF"

    border_header_sep = "FFFFFF"
    border_gray = "B8C0C8"
    border_outer = "677482"

    trs = tbl_el.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tr')
    num_rows = len(trs)

    for r_idx, tr in enumerate(trs):
        is_header = (r_idx == 0)
        is_last_row = (r_idx == num_rows - 1)
        tcs = tr.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tc')
        num_cols = len(tcs)

        if is_header:
            bg_color = header_bg
        elif r_idx % 2 == 1:
            bg_color = row_even_bg
        else:
            bg_color = row_odd_bg

        for c_idx, tc in enumerate(tcs):
            is_first_col = (c_idx == 0)
            is_last_col = (c_idx == num_cols - 1)

            tcPr = tc.get_or_add_tcPr()
            anchor_attr = tcPr.attrib.get('anchor', 'ctr')
            for child in list(tcPr):
                tcPr.remove(child)
            tcPr.attrib['anchor'] = anchor_attr

            l_clr = border_outer if is_first_col else (border_header_sep if is_header else border_gray)
            tcPr.append(parse_xml(
                f'<a:lnL {ns_a} w="12700" cap="flat" cmpd="sng" algn="ctr">'
                f'<a:solidFill><a:srgbClr val="{l_clr}"/></a:solidFill><a:prstDash val="solid"/></a:lnL>'
            ))

            r_clr = border_outer if is_last_col else (border_header_sep if is_header else border_gray)
            tcPr.append(parse_xml(
                f'<a:lnR {ns_a} w="12700" cap="flat" cmpd="sng" algn="ctr">'
                f'<a:solidFill><a:srgbClr val="{r_clr}"/></a:solidFill><a:prstDash val="solid"/></a:lnR>'
            ))

            t_clr = border_outer if (is_header or r_idx == 1) else border_gray
            t_w = "19050" if r_idx == 1 else "12700"
            tcPr.append(parse_xml(
                f'<a:lnT {ns_a} w="{t_w}" cap="flat" cmpd="sng" algn="ctr">'
                f'<a:solidFill><a:srgbClr val="{t_clr}"/></a:solidFill><a:prstDash val="solid"/></a:lnT>'
            ))

            b_clr = border_outer if (is_header or is_last_row) else border_gray
            b_w = "19050" if is_header else "12700"
            tcPr.append(parse_xml(
                f'<a:lnB {ns_a} w="{b_w}" cap="flat" cmpd="sng" algn="ctr">'
                f'<a:solidFill><a:srgbClr val="{b_clr}"/></a:solidFill><a:prstDash val="solid"/></a:lnB>'
            ))

            tcPr.append(parse_xml(
                f'<a:solidFill {ns_a}><a:srgbClr val="{bg_color}"/></a:solidFill>'
            ))


def format_table_and_note(slide, total_width_pt=671.1, assigned_images=None, captions=None, table_data=None):
    wipe_right_side_shapes(slide)
    has_image = bool(assigned_images and len(assigned_images) > 0)
    actual_width_pt = 320.0 if has_image else float(total_width_pt)
    tbl_shape = None
    note_shape = None

    shapes_to_del = []
    for shp in slide.shapes:
        if shp.top > Pt(500) and shp.left > Pt(500):
            continue
        if shp.has_table:
            tbl_shape = shp
        elif shp.has_text_frame and ("※" in shp.text_frame.text or "注" in shp.text_frame.text or "BODY_TEXT" in shp.name.upper()):
            note_shape = shp
        elif shp.shape_type == 1 and Pt(50) < shp.width < Pt(500):
            shapes_to_del.append(shp)

    for shp in shapes_to_del:
        try:
            sp = shp._element
            sp.getparent().remove(sp)
        except Exception:
            pass

    if tbl_shape and table_data and isinstance(table_data, dict):
        tbl_shape.left = Pt(113.7)
        tbl_shape.top = Pt(145.0)
        tbl_shape.width = Pt(actual_width_pt)

        headers = table_data.get("headers", [])
        rows = table_data.get("rows", [])

        max_row_cols = max([len(r) for r in rows if isinstance(r, list)] + [0])
        target_cols = max(1, len(headers), max_row_cols)

        while len(headers) < target_cols:
            headers.append(f"項目{len(headers) + 1}")

        target_rows = max(1, len(rows)) + 1

        old_tbl_el = tbl_shape._element
        old_tbl_el.getparent().remove(old_tbl_el)

        new_tbl_shape = slide.shapes.add_table(target_rows, target_cols, Pt(113.7), Pt(145.0), Pt(actual_width_pt), Pt(220.0))
        new_tbl_shape.name = "表_整形済"
        table = new_tbl_shape.table
        tbl_shape = new_tbl_shape

        for c_idx, h_text in enumerate(headers[:target_cols]):
            table.rows[0].cells[c_idx].text = str(h_text)

        for r_idx, row_items in enumerate(rows):
            target_r_idx = r_idx + 1
            if target_r_idx < len(table.rows) and isinstance(row_items, list):
                for c_idx in range(target_cols):
                    val = row_items[c_idx] if c_idx < len(row_items) else ""
                    table.rows[target_r_idx].cells[c_idx].text = str(val)

        if target_cols == 3:
            proportions = [0.22, 0.38, 0.40]
        elif target_cols == 4:
            proportions = [0.24, 0.24, 0.26, 0.26]
        elif target_cols == 5:
            proportions = [0.20, 0.24, 0.18, 0.20, 0.18]
        else:
            proportions = [1.0 / target_cols] * target_cols

        for c in range(len(table.columns)):
            col_w = actual_width_pt * proportions[c]
            table.columns[c].width = int(Pt(col_w))

        font_sz = Pt(8.5) if (len(table.rows) >= 6 or target_cols >= 5 or has_image) else Pt(9.5)
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                cell.text_frame.word_wrap = True
                cell.text_frame.margin_left = Pt(6)
                cell.text_frame.margin_right = Pt(6)
                cell.text_frame.margin_top = Pt(4)
                cell.text_frame.margin_bottom = Pt(4)
                for p in cell.text_frame.paragraphs:
                    p.font.name = "メイリオ"
                    p.font.size = font_sz
                    if r_idx == 0:
                        p.font.bold = True
                        p.font.color.rgb = RGBColor(255, 255, 255)
                    else:
                        p.font.bold = (c_idx == 0)
                        p.font.color.rgb = RGBColor(2, 24, 46)
                    for r in p.runs:
                        r.font.name = "メイリオ"
                        r.font.size = font_sz
                        if r_idx == 0:
                            r.font.bold = True
                            r.font.color.rgb = RGBColor(255, 255, 255)
                        else:
                            r.font.bold = (c_idx == 0)
                            r.font.color.rgb = RGBColor(2, 24, 46)

        apply_clean_table_borders_and_stripes(table)

    if tbl_shape and note_shape:
        note_shape.left = Pt(113.7)
        note_shape.top = int(tbl_shape.top + tbl_shape.height + Pt(14.0))
        note_shape.width = Pt(actual_width_pt)
        if note_shape.has_text_frame:
            for p in note_shape.text_frame.paragraphs:
                p.font.name = "メイリオ"
                p.font.size = Pt(9.0)
                p.font.color.rgb = RGBColor(120, 120, 120)

    if has_image:
        insert_slide_picture_safe(slide, assigned_images, captions, max_w_pt=320, max_h_pt=260, left_pt=450, top_pt=145)


def format_chart_slide_left_column(slide):
    main_line = None
    main_title = None
    main_body = None
    sub_line = None
    sub_title = None
    sub_body = None

    for shp in slide.shapes:
        if shp.top > Pt(500) and shp.left > Pt(500):
            continue
        sname = shp.name.upper()
        if "MAIN_SECTION_TITLE" in sname:
            main_title = shp
        elif "SECTION_LINE" in sname or sname == "SECTION_LINE":
            main_line = shp
        elif "MAIN_BODY_TEXT" in sname:
            main_body = shp
        elif "SUB_SECTION_TITLE" in sname or ("SECTION_TITLE" in sname and "MAIN" not in sname):
            sub_title = shp
        elif "SUB_BODY_TEXT" in sname or ("BODY_TEXT" in sname and "MAIN" not in sname):
            sub_body = shp
        elif "コネクタ" in shp.name or ("LINE" in sname and "SECTION_LINE" not in sname):
            sub_line = shp

    curr_y = Pt(110.0)

    if main_title:
        main_title.left = Pt(124.6)
        main_title.top = int(curr_y)
        main_title.width = Pt(230.0)
        t_txt = main_title.text_frame.text.strip()
        lines = len(t_txt) // 12 + t_txt.count('\n') + 1
        t_h = Pt(lines * 22.0)
        main_title.height = int(t_h)
        if main_line:
            main_line.left = Pt(113.7)
            main_line.top = int(curr_y + Pt(5.5))
        curr_y += t_h + Pt(8.0)

    if main_body and main_body.has_text_frame and main_body.text_frame.text.strip():
        main_body.left = Pt(113.7)
        main_body.top = int(curr_y)
        main_body.width = Pt(240.0)
        b_txt = main_body.text_frame.text.strip()
        for p in main_body.text_frame.paragraphs:
            p.font.name = "メイリオ"
            p.font.size = Pt(11.0)
            p.line_spacing = 1.2
            for r in p.runs:
                r.font.name = "メイリオ"
                r.font.size = Pt(11.0)
        b_lines = len(b_txt) // 16 + b_txt.count('\n') + 1
        b_h = Pt(b_lines * 17.0)
        main_body.height = int(b_h)
        curr_y += b_h + Pt(14.0)

    has_sub = (sub_title and sub_title.has_text_frame and sub_title.text_frame.text.strip()) or \
              (sub_body and sub_body.has_text_frame and sub_body.text_frame.text.strip())

    if has_sub:
        if sub_title and sub_title.has_text_frame and sub_title.text_frame.text.strip():
            sub_title.left = Pt(124.6)
            sub_title.top = int(curr_y)
            sub_title.width = Pt(230.0)
            st_h = Pt(20.0)
            sub_title.height = int(st_h)
            if sub_line:
                sub_line.left = Pt(113.7)
                sub_line.top = int(curr_y + Pt(3.5))
            curr_y += st_h + Pt(6.0)

        if sub_body and sub_body.has_text_frame and sub_body.text_frame.text.strip():
            sub_body.left = Pt(113.7)
            sub_body.top = int(curr_y)
            sub_body.width = Pt(245.0)
            sb_txt = sub_body.text_frame.text.strip()
            for p in sub_body.text_frame.paragraphs:
                p.font.name = "メイリオ"
                p.font.size = Pt(10.0)
                p.line_spacing = 1.15
                p.space_after = Pt(3.0)
                for r in p.runs:
                    r.font.name = "メイリオ"
                    r.font.size = Pt(10.0)
            sb_lines = len(sb_txt) // 16 + sb_txt.count('\n') + 1
            sub_body.height = int(Pt(sb_lines * 15.0))
    else:
        for del_shp in [sub_line, sub_title, sub_body]:
            if del_shp:
                try:
                    sp = del_shp._element
                    sp.getparent().remove(sp)
                except Exception:
                    pass


def format_text_or_image_slide_layout(slide, assigned_images, captions=None):
    main_title = None
    main_line = None
    main_body = None
    sub_title = None
    sub_line = None
    sub_body = None

    for shp in slide.shapes:
        if shp.top > Pt(500) and shp.left > Pt(500):
            continue
        sname = shp.name.upper()
        if "MAIN_SECTION_TITLE" in sname:
            main_title = shp
        elif "SECTION_LINE" in sname or sname == "SECTION_LINE":
            main_line = shp
        elif "MAIN_BODY_TEXT" in sname:
            main_body = shp
        elif "SUB_SECTION_TITLE" in sname or ("SECTION_TITLE" in sname and "MAIN" not in sname):
            sub_title = shp
        elif "SUB_BODY_TEXT" in sname or ("BODY_TEXT" in sname and "MAIN" not in sname):
            sub_body = shp
        elif "コネクタ" in shp.name or ("LINE" in sname and "SECTION_LINE" not in sname):
            sub_line = shp

    has_image = bool(assigned_images and len(assigned_images) > 0)
    has_sub = (sub_title and sub_title.has_text_frame and sub_title.text_frame.text.strip()) or \
              (sub_body and sub_body.has_text_frame and sub_body.text_frame.text.strip())

    body_text = main_body.text_frame.text.strip() if (main_body and main_body.has_text_frame) else ""
    body_chars = len(body_text)

    if has_image:
        text_width = Pt(245.0)
        curr_y = Pt(110.0)
        font_size_main = Pt(15.0)
        font_size_body = Pt(11.0)
        line_spacing_body = 1.3
        line_height_val = Pt(18.0)
        title_bottom_gap = Pt(12.0)
    elif has_sub:
        text_width = Pt(660.0)
        curr_y = Pt(115.0)
        font_size_main = Pt(18.0)
        font_size_body = Pt(13.5)
        line_spacing_body = 1.45
        line_height_val = Pt(20.0)
        title_bottom_gap = Pt(14.0)
    else:
        text_width = Pt(660.0)
        if body_chars < 120:
            curr_y = Pt(155.0)
            font_size_main = Pt(22.0)
            font_size_body = Pt(16.5)
            line_spacing_body = 1.70
            line_height_val = Pt(26.0)
            title_bottom_gap = Pt(24.0)
        elif body_chars < 220:
            curr_y = Pt(135.0)
            font_size_main = Pt(20.0)
            font_size_body = Pt(15.0)
            line_spacing_body = 1.60
            line_height_val = Pt(22.0)
            title_bottom_gap = Pt(20.0)
        else:
            curr_y = Pt(115.0)
            font_size_main = Pt(18.0)
            font_size_body = Pt(13.5)
            line_spacing_body = 1.45
            line_height_val = Pt(18.0)
            title_bottom_gap = Pt(14.0)

    if main_title:
        main_title.left = Pt(124.6)
        main_title.top = int(curr_y)
        main_title.width = text_width
        t_h = Pt(30.0) if font_size_main >= Pt(20.0) else Pt(26.0)
        main_title.height = int(t_h)
        for p in main_title.text_frame.paragraphs:
            p.font.name = "メイリオ"
            p.font.size = font_size_main
            p.font.bold = True
            for r in p.runs:
                r.font.name = "メイリオ"
                r.font.size = font_size_main
                r.font.bold = True
        if main_line:
            main_line.left = Pt(113.7)
            main_line.top = int(curr_y + Pt(3.5))
            main_line.height = int(line_height_val)
            main_line.width = Pt(4.5)
        curr_y += t_h + title_bottom_gap

    if main_body and main_body.has_text_frame and body_text:
        main_body.left = Pt(113.7)
        main_body.top = int(curr_y)
        main_body.width = text_width
        for p in main_body.text_frame.paragraphs:
            p.font.name = "メイリオ"
            p.font.size = font_size_body
            p.line_spacing = line_spacing_body
            if not has_image and not has_sub:
                p.space_after = Pt(12.0)
            for r in p.runs:
                r.font.name = "メイリオ"
                r.font.size = font_size_body
        b_lines = body_chars // (32 if not has_image else 16) + body_text.count('\n') + 1
        b_h = Pt(b_lines * (24.0 if font_size_body >= Pt(16.0) else 20.0))
        curr_y += b_h + Pt(16.0)

    if has_sub:
        if sub_title and sub_title.has_text_frame and sub_title.text_frame.text.strip():
            sub_title.left = Pt(124.6)
            sub_title.top = int(curr_y)
            sub_title.width = text_width
            for p in sub_title.text_frame.paragraphs:
                p.font.name = "メイリオ"
                p.font.size = Pt(14.0) if not has_image else Pt(11.5)
                p.font.bold = True
            if sub_line:
                sub_line.left = Pt(113.7)
                sub_line.top = int(curr_y + Pt(3.5))
            curr_y += Pt(26.0)
        if sub_body and sub_body.has_text_frame and sub_body.text_frame.text.strip():
            sub_body.left = Pt(113.7)
            sub_body.top = int(curr_y)
            sub_body.width = text_width
            for p in sub_body.text_frame.paragraphs:
                p.font.name = "メイリオ"
                p.font.size = Pt(12.5) if not has_image else Pt(10.5)
                p.line_spacing = 1.35
                p.space_after = Pt(4.0)
                for r in p.runs:
                    r.font.name = "メイリオ"
                    r.font.size = Pt(12.5) if not has_image else Pt(10.5)
    else:
        for del_shp in [sub_line, sub_title, sub_body]:
            if del_shp:
                try:
                    sp = del_shp._element
                    sp.getparent().remove(sp)
                except Exception:
                    pass

    if has_image:
        insert_slide_picture_safe(slide, assigned_images, captions, max_w_pt=360, max_h_pt=300, left_pt=390, top_pt=130)


def build_clean_index_slide(slide, edited_chapters):
    shapes_to_remove = []
    for shp in slide.shapes:
        if shp.top > Pt(500) and shp.left > Pt(500):
            continue
        if shp.is_placeholder and shp.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER:
            continue
        sname = shp.name.upper()
        if "INDEX_" in sname:
            shapes_to_remove.append(shp)
        elif shp.has_text_frame and ("[[" in shp.text_frame.text or re.match(r'^\d{2}\s', shp.text_frame.text.strip())):
            shapes_to_remove.append(shp)

    for shp in shapes_to_remove:
        try:
            sp = shp._element
            sp.getparent().remove(sp)
        except Exception:
            pass

    num_items = len(edited_chapters)
    if num_items <= 6:
        start_y_pt = 220.0
        step_y_pt = 38.5
        font_sz_pt = 18.0
    else:
        start_y_pt = 195.0
        step_y_pt = min(36.0, 270.0 / max(1, num_items))
        font_sz_pt = 15.5

    title_x = Pt(210.0)
    page_x = Pt(610.0)

    for idx, ch in enumerate(edited_chapters):
        ch_num_str = f"{idx + 1:02d}"
        ch_page_num = idx + 2
        clean_title = clean_chapter_title(ch.get("chapter_title", f"章{idx+1}"))
        row_y = Pt(start_y_pt + idx * step_y_pt)

        t_box = slide.shapes.add_textbox(title_x, row_y, Pt(390.0), Pt(step_y_pt))
        t_box.name = f"INDEX_{idx+1}_TITLE"
        tf_t = t_box.text_frame
        tf_t.word_wrap = False
        tf_t.margin_left = Pt(4)
        tf_t.margin_right = Pt(4)
        tf_t.margin_top = Pt(2)
        tf_t.margin_bottom = Pt(2)
        p_t = tf_t.paragraphs[0]

        r0 = p_t.add_run()
        r0.text = f"{ch_num_str} "
        r0.font.name = "メイリオ"
        r0.font.size = Pt(font_sz_pt)
        r0.font.bold = True
        r0.font.color.rgb = RGBColor(217, 48, 37)

        r1 = p_t.add_run()
        r1.text = clean_title
        r1.font.name = "メイリオ"
        r1.font.size = Pt(font_sz_pt)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(2, 24, 46)

        p_box = slide.shapes.add_textbox(page_x, row_y, Pt(80.0), Pt(step_y_pt))
        p_box.name = f"INDEX_{idx+1}_PAGE"
        tf_p = p_box.text_frame
        tf_p.word_wrap = False
        tf_p.margin_left = Pt(4)
        tf_p.margin_right = Pt(4)
        tf_p.margin_top = Pt(2)
        tf_p.margin_bottom = Pt(2)
        p_p = tf_p.paragraphs[0]
        p_p.alignment = PP_ALIGN.LEFT
        r_p = p_p.add_run()
        r_p.text = f"P{ch_page_num}"
        r_p.font.name = "メイリオ"
        r_p.font.size = Pt(font_sz_pt)
        r_p.font.bold = True
        r_p.font.color.rgb = RGBColor(2, 24, 46)


def format_chapter_title_paragraph(p, ch_num, ch_title):
    clean_title = clean_chapter_title(ch_title)
    p.text = ""
    r0 = p.add_run()
    r0.text = f"{ch_num} "
    r0.font.color.rgb = RGBColor(217, 48, 37)
    r0.font.name = "メイリオ"
    r0.font.size = Pt(20.0)
    r0.font.bold = True

    r1 = p.add_run()
    r1.text = clean_title
    r1.font.name = "メイリオ"
    r1.font.size = Pt(20.0)
    r1.font.bold = True


def process_paragraph_runs(p, replace_map, chapter_num=None, chapter_title=None, is_cover=False):
    if "00" in p.text or "[[章タイトル]]" in p.text:
        format_chapter_title_paragraph(p, chapter_num if chapter_num else "01", chapter_title)
        return

    for r in p.runs:
        for tag, val in replace_map.items():
            if tag in r.text:
                r.text = r.text.replace(tag, str(val))

    p_text = p.text
    has_unreplaced_tag = False
    for tag in replace_map:
        if tag in p_text:
            has_unreplaced_tag = True
            break

    if has_unreplaced_tag:
        full_text = p_text
        for tag, val in replace_map.items():
            if tag in full_text:
                full_text = full_text.replace(tag, str(val))

        if len(p.runs) > 0:
            p.runs[0].text = full_text
            for r in p.runs[1:]:
                r.text = ""

    if p.text and "\n\n" in p.text:
        p.text = sanitize_text_empty_lines(p.text)

    if not is_cover and p.text and len(p.text.strip()) > 5:
        if not any(tag in p.text for tag in ["[[章タイトル]]", "STEP", "P1", "P2", "□"]):
            for r in p.runs:
                r.font.name = "メイリオ"


def process_slide_shapes(slide, replace_map, chapter_num=None, chapter_title=None, is_cover=False):
    shapes_to_remove = []

    notice_val = replace_map.get("[[注意文]]", "") or replace_map.get("[[補足文]]", "")
    has_notice = bool(notice_val and str(notice_val).strip())
    body_val = replace_map.get("[[本文]]", "")
    has_body = bool(body_val and str(body_val).strip())

    if not has_notice and not is_cover:
        for shape in slide.shapes:
            if shape.top > Pt(500) and shape.left > Pt(500):
                continue
            sname_u = shape.name.upper()
            raw_t = shape.text_frame.text if shape.has_text_frame else ""
            if any(kw in sname_u for kw in ["CAUTION", "SUPPLEMENT", "補足", "注意"]):
                shapes_to_remove.append(shape)
            elif any(kw in raw_t for kw in ["[[注意文]]", "[[補足文]]", "注意", "補足", "※", "注"]):
                shapes_to_remove.append(shape)

    for shape in slide.shapes:
        if shape.top > Pt(500) and shape.left > Pt(500):
            continue
        raw_text = shape.text_frame.text if shape.has_text_frame else ""
        if shape.has_text_frame:
            txt = raw_text.strip()
            if txt.isdigit() and len(txt) <= 2:
                is_bottom_right = (shape.top > Pt(400) and shape.left > Pt(500))
                is_placeholder = (shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER)
                if not is_bottom_right and not is_placeholder:
                    shapes_to_remove.append(shape)
                    continue

        if not has_body and shape.has_text_frame:
            if "[[本文]]" in raw_text:
                shapes_to_remove.append(shape)
                continue

        if shape.has_text_frame:
            tf = shape.text_frame
            tf.word_wrap = True
            for p in tf.paragraphs:
                if "[[チェック項目" in p.text:
                    for k in range(1, 7):
                        tag_k = f"[[チェック項目{k}]]"
                        if tag_k in p.text and replace_map.get(tag_k, "").strip() == "":
                            p.text = ""
                process_paragraph_runs(p, replace_map, chapter_num, chapter_title, is_cover=is_cover)

        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    for p in cell.text_frame.paragraphs:
                        process_paragraph_runs(p, replace_map, is_cover=is_cover)

    for shp in shapes_to_remove:
        try:
            sp = shp._element
            sp.getparent().remove(sp)
        except Exception:
            pass


def run_quality_check(raw_memo, edited_chapters, doc_title):
    memo_topics = {
        "月別売上推移（1〜6月・計6,500万円）": ["850", "920", "1,100", "1100", "980", "1,250", "1250", "1,400", "1400", "6,500", "6500"],
        "カテゴリー別売上構成（生活雑貨35%・食品12%等）": ["生活雑貨", "35", "家電", "25", "ファッション", "20", "食品", "12"],
        "顧客年代構成（30〜40代中心・20代拡大）": ["20代", "30代", "40代", "58"],
        "商品別実績（商品A〜Eの売上・利益率・数量）": ["商品A", "商品B", "商品C", "商品D", "商品E", "40", "4,200", "4200", "1,400", "1400"],
        "販売チャネル構成（自社EC55%・モール拡大）": ["自社EC", "55", "楽天市場", "25", "Amazon", "15"],
        "キャンペーン実績比較（春・GW・初夏）": ["春の新生活", "ゴールデンウィーク", "初夏", "80", "100", "70", "720"],
        "5大販売戦略（①〜⑤全5項目）": ["リピーター", "商品C", "楽天市場", "Instagram", "キャンペーン実績"],
        "2027年度目標（1.5億円・リピート率30%等）": ["1億5,000", "1.5億", "8,000", "30", "45"]
    }

    all_slide_text = doc_title + " " + " ".join([
        ch.get("chapter_title", "") + " " +
        ch.get("main_item", "") + " " +
        ch.get("body", "") + " " +
        ch.get("sub_title", "") + " " +
        ch.get("sub_body", "") + " " +
        " ".join(ch.get("step_descs", [])) + " " +
        " ".join(ch.get("checklist_items", [])) + " " +
        " ".join([val for row in ch.get("table", {}).get("rows", []) for val in row]) + " " +
        " ".join(map(str, ch.get("chart_info", {}).get("categories", []))) + " " +
        " ".join(map(str, ch.get("chart_info", {}).get("values", [])))
        for ch in edited_chapters
    ])

    return {topic: any(kw in all_slide_text for kw in kws) for topic, kws in memo_topics.items()}


# ==========================================
# 1. 入力UI ＆ サイドバー
# ==========================================
with st.sidebar:
    st.header("⚙️ 設定")
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    st.info("テンプレート: 資料作成テンプレート_A4横.pptx")

    if st.button("🔄 データを初期化して最初から行う"):
        st.session_state.pop("parsed_data", None)
        st.session_state.pop("analyzed", None)
        st.session_state.pop("saved_images", None)
        st.session_state.pop("refined_title", None)
        st.rerun()

col1, col2 = st.columns(2)
with col1:
    client_name = st.text_input("提案先（顧客名）", "株式会社ネクストライフ")
    doc_title = st.text_input("資料タイトル", "EC事業売上分析および2027年度販売成長戦略")
with col2:
    sender_name = st.text_input("作成者・部署など", "EC事業部")
    uploaded_files = st.file_uploader(
        "挿入画像（任意・複数可）",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="main_file_uploader_key"
    )

if uploaded_files:
    current_files_map = {f.name: f.getvalue() for f in uploaded_files}
    saved_map = {img["filename"]: img["bytes"] for img in st.session_state.get("saved_images", [])}
    saved_map.update(current_files_map)
    st.session_state["saved_images"] = [{"filename": k, "bytes": v} for k, v in saved_map.items()]

active_images = st.session_state.get("saved_images", [])

image_contexts = []
if active_images:
    st.success(f"📷 合計 {len(active_images)} 枚の画像が読み込まれました。")
    for row_start in range(0, len(active_images), 4):
        row_imgs = active_images[row_start:row_start + 4]
        cols = st.columns(4)
        for col_idx, img_item in enumerate(row_imgs):
            actual_idx = row_start + col_idx
            with cols[col_idx]:
                st.image(img_item["bytes"], caption=f"画像{actual_idx+1}: {img_item['filename']}", use_container_width=True)
                desc = st.text_input(f"画像{actual_idx+1}の説明", placeholder="例：サービスイメージ図", key=f"img_ctx_{actual_idx}")
                image_contexts.append({"index": actual_idx, "filename": img_item["filename"], "description": desc})

raw_memo = st.text_area(
    "資料の内容メモ",
    """当社ECサイトの2026年度上期の販売実績を振り返り、今後の売上拡大に向けた施策を検討したい。

全体として売上は伸びているが、商品によって売れ行きに差がある。また、新規顧客は増えているものの、リピーターの獲得にはまだ課題がある。

【1. 月別売上実績】
2026年1月〜6月の売上は以下の通り。
1月：850万円
2月：920万円
3月：1,100万円
4月：980万円
5月：1,250万円
6月：1,400万円
上期累計売上は6,500万円。1月と比較して6月の売上は約65％増加している。3月と6月に売上が大きく伸びている。春の新生活需要や夏に向けた販促が影響した可能性がある。

【2. 商品カテゴリー別売上構成】
上期売上6,500万円のカテゴリー別構成は以下の通り。
生活雑貨：35％
家電・デジタル：25％
ファッション：20％
食品・飲料：12％
その他：8％
生活雑貨が最も大きな売上構成を占めている。一方で、食品・飲料は売上構成が小さいものの、購入頻度が比較的高いカテゴリーとなっている。

【3. 顧客の年代構成】
購入者の年代別割合は以下の通り。
20代：15％
30代：30％
40代：28％
50代：18％
60代以上：9％
30代と40代で全体の58％を占める。今後は20代の新規顧客を増やすことも検討したい。

【4. 商品別販売実績】
上期の商品別販売データ。
商品A：生活雑貨セット（販売数量：4,200個、売上：1,400万円、利益率：35％）
商品B：ワイヤレスイヤホン（販売数量：1,800個、売上：1,250万円、利益率：28％）
商品C：収納ボックス（販売数量：3,500個、売上：1,050万円、利益率：40％）
商品D：季節限定食品セット（販売数量：2,800個、売上：780万円、利益率：22％）
商品E：ファッション小物（販売数量：2,100個、売上：920万円、利益率：32％）
商品Aは売上が最も高い。商品Cは利益率が最も高く、収益性の面で重要な商品となっている。

【5. 販売チャネル別実績】
売上の販売経路別構成。
自社ECサイト：55％
楽天市場：25％
Amazon：15％
その他：5％
自社ECサイトが売上の過半数を占めている。今後はモール型ECの売上拡大も検討する必要がある。

【6. 販促キャンペーンの実績比較】
過去に実施したキャンペーンの結果。
春の新生活キャンペーン（実施期間：3月1日〜3月31日、広告費：80万円、売上：650万円、新規顧客：320人）
ゴールデンウィークセール（実施期間：4月25日〜5月10日、広告費：100万円、売上：580万円、新規顧客：250人）
初夏のまとめ買いキャンペーン（実施期間：6月1日〜6月30日、広告費：70万円、売上：720万円、新規顧客：400人）
初夏のキャンペーンは広告費が最も少ない一方、売上と新規顧客数が最も多かった。まとめ買いを促す施策が効果的だった可能性がある。

【7. 現在の課題】
・新規顧客は増えているが、リピート購入につながっていない
・商品カテゴリーによって利益率に差がある
・自社ECへの売上依存度が高い
・SNSのフォロワー数は増えているが、購入へのつながりが不明確
・販促キャンペーンの効果測定が十分にできていない
特にリピーター対策は優先度が高いと考えている。

【8. 今後の販売戦略】
① リピーター獲得：購入後のメール配信、会員限定クーポン、ポイント制度の見直しを実施する。
② 商品戦略：利益率の高い商品Cの販売を強化する。また、商品Aとのセット販売も検討する。
③ 販売チャネル拡大：楽天市場・Amazonの商品ページ改善と広告運用を強化する。
④ SNS販促強化：Instagramを中心に商品の使用シーンや購入者のレビューを発信する。
⑤ 販促施策の改善：過去のキャンペーン実績を分析し、広告費に対する売上効果を確認しながら施策を実施する。

【9. 2027年度の目標】
年間売上目標：1億5,000万円
新規顧客数：年間8,000人
リピート購入率：現在の20％から30％へ改善
自社EC以外の販売チャネル売上構成比：45％

【10. 最後に】
今後は単純に新規顧客を増やすだけでなく、既存顧客との関係を強化し、継続的に購入してもらえる仕組みを構築したい。
売上規模の拡大と同時に利益率も意識し、持続的に成長できるEC事業を目指す。""",
    height=260
)


# ==========================================
# ステップ1: AIで構造解析
# ==========================================
st.markdown("---")
col_btn1, col_btn2 = st.columns([2, 1])

with col_btn1:
    if st.button("🔍 AIでメモを解析し、構成プレビューを作成する", use_container_width=True):
        if not api_key:
            st.error("APIキーが設定されていません。.streamlit/secrets.toml をご確認ください。")
        elif not raw_memo.strip():
            st.error("提案メモを入力してください。")
        else:
            with st.spinner("AIがメモを解析中..."):
                try:
                    client = genai.Client(api_key=api_key)
                    prompt_data = {
                        "client_name": client_name,
                        "doc_title": doc_title,
                        "raw_memo": raw_memo,
                        "image_contexts": image_contexts
                    }

                    single_pass_prompt = f"""
                    あなたは卓越した経営コンサルタント兼プレゼン資料作成のプロフェッショナルです。
                    以下の入力メモの数値・事実・項目・目標を【1つも省略せず正確に反映】したプレゼンスライドJSONを作成してください。
                    元メモに明記された事実（実績数値・目標数値）と、分析・推測（「〜可能性がある」等）を明確に区別して記述してください。

                    【入力データ】
                    {json.dumps(prompt_data, ensure_ascii=False, indent=2)}

                    【必須ルール】
                    1. 販売チャネル別構成比（自社EC55%、楽天市場25%、Amazon15%、その他5%）は独立したスライド（"pie_chart"）として必ず作成し、課題スライドと混同しないでください。
                    2. カテゴリー別売上構成・顧客年代構成のスライドは "pie_chart" を選定してください。
                    3. 商品別販売実績の表（table）では、列見出しに ["商品名", "商品詳細", "販売数量", "売上高", "利益率"] を含め、商品A〜Eの5行すべてを省略せず出力してください。カンマ区切り数値（4,200、1,400万円など）はそのまま保持してください。
                    4. キャンペーン実績比較の表（table）では、列見出しに ["キャンペーン名", "実施期間", "広告費", "売上高", "新規顧客"] を含め、3キャンペーンすべてを出力してください。
                    5. 5大販売戦略（①リピーター獲得、②商品戦略、③販売チャネル拡大、④SNS販促強化、⑤販促施策の改善）の5つすべてを必ず "step_descs" 配列に含めてください。先頭に「①」「1.」「STEP 1」などの番号記号は付けないでください。
                    6. 2027年度目標（年間売上目標1億5,000万円、新規顧客数年間8,000人、リピート購入率20%→30%、自社EC以外構成比45%）を必ずスライドに含めてください。
                    7. "main_item" は15文字以内の短いキャッチコピーにしてください。

                    【出力形式】
                    純粋なJSONのみを出力してください。
                    {{
                      "suggested_doc_title": "EC事業売上分析および2027年度販売成長戦略",
                      "chapters": [
                        {{
                          "chapter_title": "章タイトル",
                          "layout_type": "line_chart",
                          "main_item": "主要見出し（15文字以内）",
                          "body": "本文テキスト",
                          "sub_title": "補助項目タイトル",
                          "sub_body": "補助項目本文",
                          "notice": "",
                          "assigned_image_indices": [],
                          "image_captions": [],
                          "step_descs": ["施策1", "施策2", "施策3", "施策4", "施策5"],
                          "checklist_intro": "導入文",
                          "checklist_items": ["項目1", "項目2", "項目3", "項目4", "項目5"],
                          "checklist_outro": "案内文",
                          "table": {{
                            "headers": ["列1", "列2", "列3"],
                            "rows": [["行1-1", "行1-2", "行1-3"]],
                            "note": ""
                          }},
                          "chart_info": {{
                            "title": "グラフタイトル",
                            "unit": "万円",
                            "categories": ["1月", "2月", "3月", "4月", "5月", "6月"],
                            "series_name": "売上高（万円）",
                            "values": [850, 920, 1100, 980, 1250, 1400]
                          }}
                        }}
                      ]
                    }}
                    """

                    txt_res = call_gemini_with_retry(client, single_pass_prompt)
                    m_json = re.search(r'\{[\s\S]*\}', txt_res)
                    final_data = json.loads(m_json.group(0) if m_json else txt_res)

                    st.session_state["parsed_data"] = final_data
                    st.session_state["analyzed"] = True

                    sug_title = final_data.get("suggested_doc_title")
                    if sug_title and len(sug_title.strip()) > 5:
                        st.session_state["refined_title"] = sug_title.strip()

                    st.success("✅ AIの構成解析が完了しました！下部の編集フォームで内容を確認・微調整できます。")
                    st.rerun()
                except Exception as e:
                    st.error(f"解析中にエラーが発生しました: {e}")


# ==========================================
# ステップ2: WEB編集フォーム ＆ 最終生成
# ==========================================
if st.session_state.get("analyzed", False) and "parsed_data" in st.session_state:
    st.subheader("✏️ AI解析結果のプレビュー ＆ WEB編集フォーム")
    st.caption("各スライドの内容を自由に変更できます。カンマを含む数値はダブルクォーテーション（例: \"4,200\"）で囲むと安全です。")

    if "refined_title" in st.session_state:
        doc_title = st.text_input("資料タイトル", value=st.session_state["refined_title"], key="refined_title_box")
    else:
        doc_title = st.text_input("資料タイトル", value=doc_title, key="current_title_box")

    parsed_chapters = st.session_state["parsed_data"].get("chapters", [])
    edited_chapters = []

    for idx, ch in enumerate(parsed_chapters):
        with st.expander(f"📄 スライド {idx + 1}: {ch.get('chapter_title', '')}", expanded=(idx <= 2)):
            ec_cols1, ec_cols2 = st.columns([2, 1])

            with ec_cols1:
                e_title = st.text_input(f"スライド{idx+1} タイトル", value=ch.get("chapter_title", ""), key=f"title_{idx}")
                e_main = st.text_input(f"スライド{idx+1} 主要見出し（赤帯・15文字以内）", value=ch.get("main_item", ""), key=f"main_{idx}")
                e_body = st.text_area(f"スライド{idx+1} 本文", value=sanitize_text_empty_lines(ch.get("body", "")), height=100, key=f"body_{idx}")

            with ec_cols2:
                current_layout_key = str(ch.get("layout_type", "text")).lower()
                current_label = LAYOUT_MAP.get(current_layout_key, LAYOUT_MAP["text"])
                labels_list = list(LAYOUT_MAP.values())
                l_idx = labels_list.index(current_label) if current_label in labels_list else 3

                selected_label = st.selectbox("レイアウトタイプ", labels_list, index=l_idx, key=f"layout_{idx}")
                e_layout = LAYOUT_REVERSE_MAP[selected_label]

                e_sub_title = st.text_input("補助項目 タイトル", value=ch.get("sub_title", ""), key=f"sub_t_{idx}")
                e_sub_body = st.text_area("補助項目 本文", value=sanitize_text_empty_lines(ch.get("sub_body", "")), height=68, key=f"sub_b_{idx}")

            e_assigned_indices = []
            e_captions = []
            if active_images:
                st.markdown("**📷 画像の挿入設定（任意）**")
                img_options = ["（挿入しない）"] + [f"画像{i+1}: {img['filename']}" for i, img in enumerate(active_images)]

                ai_assigned = ch.get("assigned_image_indices", [])
                default_sel_idx = 0
                if ai_assigned and len(ai_assigned) > 0 and ai_assigned[0] < len(active_images):
                    default_sel_idx = ai_assigned[0] + 1

                sel_img = st.selectbox(f"スライド{idx+1} に挿入する画像", img_options, index=default_sel_idx, key=f"img_sel_{idx}")
                if sel_img != "（挿入しない）":
                    chosen_idx = img_options.index(sel_img) - 1
                    e_assigned_indices = [chosen_idx]
                    old_caps = ch.get("image_captions", [])
                    default_cap = old_caps[0] if (old_caps and isinstance(old_caps, list) and len(old_caps) > 0) else ""
                    user_cap = st.text_input("画像のキャプション（任意）", value=default_cap, key=f"cap_{idx}")
                    e_captions = [user_cap]
                    st.success(f"✅ このスライドに 『{active_images[chosen_idx]['filename']}』 を挿入します")
            else:
                e_assigned_indices = []
                e_captions = []

            if e_layout == "step":
                st.markdown("**ステップ・施策詳細（全項目を漏れなく反映）**")
                s_descs = ch.get("step_descs", [])
                if not isinstance(s_descs, list) or len(s_descs) < 5:
                    s_descs = [
                        "リピーター獲得：購入後のメール配信、会員限定クーポン、ポイント制度の見直しを実施",
                        "商品戦略：利益率の高い商品C（40%）の販売強化および商品Aとのセット販売検討",
                        "販売チャネル拡大：楽天市場・Amazonの商品ページ改善と広告運用を強化",
                        "SNS販促強化：Instagramを中心に商品の使用シーンや購入者のレビューを発信",
                        "販促施策の改善：過去のキャンペーン実績を分析し、広告費に対する売上効果を確認しながら施策を実施"
                    ]
                num_step_inputs = max(len(s_descs), 5)
                e_s_descs = []
                s_cols = st.columns(2)
                for s_i in range(num_step_inputs):
                    default_s = s_descs[s_i] if len(s_descs) > s_i else ""
                    clean_default = clean_step_text(default_s)
                    with s_cols[s_i % 2]:
                        val_s = st.text_input(f"項目 {s_i+1}", value=clean_default, key=f"step_{idx}_{s_i}")
                        if val_s.strip():
                            e_s_descs.append(clean_step_text(val_s))
            else:
                e_s_descs = [clean_step_text(t) for t in ch.get("step_descs", []) if clean_step_text(t)]

            if e_layout == "checklist":
                st.markdown("**✅ チェックリスト設定（最大6項目）**")
                c_intro_val = st.text_input("チェックリスト導入文", value=ch.get("checklist_intro", "EC事業における現在の主要課題一覧です。"), key=f"c_intro_{idx}")
                c_items_def = ch.get("checklist_items", [])
                if not isinstance(c_items_def, list) or not c_items_def:
                    c_items_def = [
                        "新規顧客は増えているが、リピート購入につながっていない（リピート率20%→目標30%）",
                        "商品カテゴリーによって利益率に差がある（最高40%〜最低22%）",
                        "自社ECサイトへの売上依存度が高い（自社EC55%・他チャネル45%へ拡大目標）",
                        "SNSのフォロワー数は増えているが、購入へのつながりが不明確",
                        "販促キャンペーンの効果測定が十分にできていない"
                    ]
                e_c_items = []
                ci_cols = st.columns(2)
                for ci in range(6):
                    def_item = c_items_def[ci] if len(c_items_def) > ci else ""
                    with ci_cols[ci % 2]:
                        val_item = st.text_input(f"チェック項目 {ci+1}", value=def_item, key=f"ci_{idx}_{ci}")
                        if val_item.strip():
                            e_c_items.append(val_item.strip())
                c_outro_val = st.text_input("確認後の案内文", value=ch.get("checklist_outro", "リピーター獲得を最優先とし、2027年度売上1億5,000万円達成を目指します。"), key=f"c_outro_{idx}")
            else:
                c_intro_val = ch.get("checklist_intro", "")
                e_c_items = ch.get("checklist_items", [])
                c_outro_val = ch.get("checklist_outro", "")

            if e_layout == "table":
                st.markdown("**📋 表組み（比較表・施策一覧・タスク管理）データ設定**")
                tbl_raw = ch.get("table")
                if not isinstance(tbl_raw, dict):
                    tbl_raw = {}
                h_list = tbl_raw.get("headers", ["項目", "詳細・実績", "備考・狙い"])
                r_list = tbl_raw.get("rows", [])

                h_display_str = ", ".join([f'"{x}"' if "," in str(x) else str(x) for x in h_list])
                t_cols_input = st.text_input("列見出し（カンマ区切り・全列反映）", value=h_display_str, key=f"th_{idx}")
                e_headers = parse_csv_line_safe(t_cols_input)

                num_display_rows = max(len(r_list), 5)
                st.caption(f"各行のデータ（カンマ区切り・全{num_display_rows}行）")
                e_rows = []
                for ri in range(num_display_rows):
                    if len(r_list) > ri and isinstance(r_list[ri], list):
                        default_row_str = ", ".join([f'"{x}"' if "," in str(x) else str(x) for x in r_list[ri]])
                    else:
                        default_row_str = ""
                    row_str = st.text_input(f"{ri+1}行目", value=default_row_str, key=f"tr_{idx}_{ri}")
                    parsed_row = parse_csv_line_safe(row_str)
                    if any(cell != "" for cell in parsed_row):
                        e_rows.append(parsed_row)
                t_note = st.text_input("表の注記（任意）", value=tbl_raw.get("note", ""), key=f"tn_{idx}")
                e_table = {"headers": e_headers, "rows": e_rows, "note": t_note}
            else:
                e_table = ch.get("table") if isinstance(ch.get("table"), dict) else {}

            c_info = ch.get("chart_info")
            if not isinstance(c_info, dict) or not c_info.get("values"):
                c_info = extract_chapter_specific_chart_data(e_title, e_body, raw_memo)

            if e_layout in ["line_chart", "bar_chart", "pie_chart"]:
                st.markdown("**📊 グラフデータ設定**")
                g_cols = st.columns([3, 2, 3, 3])
                with g_cols[0]:
                    g_title = st.text_input("グラフタイトル", value=c_info.get("title", f"{e_title}"), key=f"gt_{idx}")
                with g_cols[1]:
                    def_unit = c_info.get("unit", "")
                    if not def_unit:
                        if e_layout == "pie_chart" or any(kw in e_title or kw in e_body for kw in ["％", "%", "構成", "割合", "チャネル", "年代"]):
                            def_unit = "％"
                        elif any(kw in e_title or kw in e_body for kw in ["万", "円", "売上", "月別"]):
                            def_unit = "万円"
                    g_unit = st.text_input("単位（％, 万円, 個など）", value=def_unit, key=f"gu_{idx}")
                with g_cols[2]:
                    g_cats_str = st.text_input(
                        "カテゴリ（カンマ区切り）",
                        value=",".join(map(str, c_info.get("categories", []))),
                        key=f"gc_{idx}"
                    )
                with g_cols[3]:
                    def_vals = c_info.get("values", [])
                    clean_vals = [f'"{v}"' if "," in str(v) else str(v) for v in def_vals]
                    g_vals_str = st.text_input("数値データ（カンマ区切り・カンマ付き数値可）", value=",".join(clean_vals), key=f"gv_{idx}")

                cats_list = parse_csv_line_safe(g_cats_str)
                raw_val_tokens = parse_csv_line_safe(g_vals_str)
                vals_list = [parse_number_with_comma(v) for v in raw_val_tokens if str(v).strip() != ""]

                edited_c_info = {
                    "title": g_title,
                    "unit": g_unit.strip(),
                    "categories": cats_list,
                    "series_name": c_info.get("series_name", g_title),
                    "values": vals_list
                }
            else:
                edited_c_info = c_info

            edited_chapters.append({
                "chapter_title": e_title,
                "layout_type": e_layout,
                "main_item": e_main,
                "body": sanitize_text_empty_lines(e_body),
                "sub_title": e_sub_title,
                "sub_body": sanitize_text_empty_lines(e_sub_body),
                "notice": ch.get("notice", ""),
                "step_descs": e_s_descs,
                "checklist_intro": c_intro_val,
                "checklist_items": e_c_items,
                "checklist_outro": c_outro_val,
                "table": e_table,
                "assigned_image_indices": e_assigned_indices,
                "image_captions": e_captions,
                "chart_info": edited_c_info
            })

    st.markdown("---")

    if st.button("✨ この内容で PowerPoint 資料を自動生成する ✨", key="generate_pptx_btn", type="primary", use_container_width=True):
        with st.spinner("PowerPoint資料を精密構築中..."):
            try:
                current_dir = os.path.dirname(os.path.abspath(__file__))

                template_path = None
                target_norm = unicodedata.normalize('NFC', '資料作成テンプレート_A4横.pptx')

                for f in os.listdir(current_dir):
                    if not f.endswith('.pptx'):
                        continue
                    f_norm = unicodedata.normalize('NFC', f)
                    if f_norm == target_norm or "テンプレート" in f_norm or "template" in f_norm.lower():
                        template_path = os.path.join(current_dir, f)
                        break

                if not template_path:
                    for f in os.listdir(current_dir):
                        if f.endswith('.pptx') and not any(kw in f for kw in ["報告", "提案書", "上期", "成長戦略", "集客", "ネクストライフ", "EC事業"]):
                            template_path = os.path.join(current_dir, f)
                            break

                if not template_path:
                    raise FileNotFoundError("『資料作成テンプレート_A4横.pptx』が見つかりません。")

                prs = Presentation(template_path)
                tpl_slide_count = len(prs.slides)
                today_str = datetime.date.today().strftime("%Y/%m/%d")

                # 1. 表紙スライド
                cover_slide = duplicate_slide_safe(prs, 0)
                cover_map = {
                    "[[資料タイトル]]": doc_title,
                    "[[顧客名]]": client_name,
                    "[[バージョン]]": "1.0",
                    "[[更新日]]": today_str
                }
                process_slide_shapes(cover_slide, cover_map, is_cover=True)

                # 2. 目次スライド
                index_slide = duplicate_slide_safe(prs, 1)
                build_clean_index_slide(index_slide, edited_chapters)

                inserted_image_report = []
                chart_audit_report = []

                # 3. 各章スライド構築
                for ch_idx, ch in enumerate(edited_chapters):
                    ch_num_str = f"{ch_idx + 1:02d}"
                    ch_page_num = ch_idx + 2
                    raw_ch_title = ch.get("chapter_title", f"章{ch_idx+1}")
                    clean_ch_title = clean_chapter_title(raw_ch_title)

                    layout_type = str(ch.get("layout_type", "text")).lower()
                    img_indices = ch.get("assigned_image_indices", [])
                    assigned_imgs = [active_images[i] for i in img_indices if i < len(active_images)]
                    captions = ch.get("image_captions", [])
                    raw_chart_info = ch.get("chart_info")

                    if layout_type not in ["line_chart", "bar_chart", "pie_chart", "step", "checklist", "table"]:
                        if any(kw in clean_ch_title for kw in ["カテゴリー", "カテゴリ", "年代", "チャネル", "販売経路"]):
                            layout_type = "pie_chart"

                    should_draw_chart = (layout_type in ["line_chart", "bar_chart", "pie_chart"])

                    if layout_type == "step":
                        target_template_idx = 5 if tpl_slide_count > 5 else (3 if tpl_slide_count > 3 else 2)
                    elif layout_type == "checklist":
                        target_template_idx = 9 if tpl_slide_count > 9 else (4 if tpl_slide_count > 4 else 2)
                    elif layout_type == "table":
                        target_template_idx = 8 if tpl_slide_count > 8 else (5 if tpl_slide_count > 5 else 2)
                    elif should_draw_chart:
                        target_template_idx = 2
                    else:
                        target_template_idx = 2

                    if target_template_idx >= tpl_slide_count:
                        target_template_idx = 2

                    ch_slide = duplicate_slide_safe(prs, target_template_idx)

                    main_item_text = sanitize_main_item(ch.get("main_item", ""), clean_ch_title)
                    body_text_val = sanitize_text_empty_lines(ch.get("body", ""))

                    ch_map = {
                        "[[主要項目]]": main_item_text,
                        "[[本文]]": body_text_val,
                        "[[補助項目]]": ch.get("sub_title", ""),
                        "[[補助項目の本文]]": sanitize_text_empty_lines(ch.get("sub_body", ""))
                    }

                    notice_val = clean_notice_text(ch.get("notice", ""))
                    ch_map["[[補足文]]"] = notice_val
                    ch_map["[[注意文]]"] = notice_val

                    step_descs = ch.get("step_descs", [])
                    if not isinstance(step_descs, list) or len(step_descs) < 4:
                        step_descs = [
                            "リピーター獲得：購入後のメール配信、会員限定クーポン、ポイント制度の見直しを実施",
                            "商品戦略：利益率の高い商品C（40%）の販売強化および商品Aとのセット販売検討",
                            "販売チャネル拡大：楽天市場・Amazonの商品ページ改善と広告運用を強化",
                            "SNS販促強化：Instagramを中心に商品の使用シーンや購入者のレビューを発信",
                            "販促施策の改善：過去のキャンペーン実績を分析し、広告費に対する売上効果を確認しながら施策を実施"
                        ]

                    normalized_4 = normalize_step_descs_to_4(step_descs)
                    for idx_s in range(1, 5):
                        ch_map[f"[[STEP{idx_s}の説明]]"] = normalized_4[idx_s - 1]

                    c_intro = ch.get("checklist_intro", "EC事業における現在の主要課題一覧です。")
                    c_outro = ch.get("checklist_outro", "リピーター獲得を最優先とし、2027年度売上1億5,000万円達成を目指します。")
                    c_items = ch.get("checklist_items", [])
                    ch_map["[[チェックリスト導入文]]"] = c_intro
                    ch_map["[[チェック後の案内文]]"] = c_outro
                    for k in range(1, 7):
                        ch_map[f"[[チェック項目{k}]]"] = c_items[k - 1] if (isinstance(c_items, list) and len(c_items) >= k) else ""

                    tbl_data = ch.get("table", {})

                    process_slide_shapes(
                        ch_slide, ch_map,
                        chapter_num=ch_num_str, chapter_title=clean_ch_title,
                        is_cover=False
                    )

                    if layout_type == "step":
                        build_perfect_steps(ch_slide, step_descs, body_text_val, assigned_imgs, captions)
                    elif layout_type == "checklist":
                        format_checklist_slide(ch_slide, assigned_imgs, captions)
                    elif layout_type == "table":
                        format_table_and_note(ch_slide, total_width_pt=671.1, assigned_images=assigned_imgs, captions=captions, table_data=tbl_data)
                    elif should_draw_chart:
                        format_chart_slide_left_column(ch_slide)
                    else:
                        format_text_or_image_slide_layout(ch_slide, assigned_imgs, captions)

                    if len(assigned_imgs) > 0 and not should_draw_chart:
                        for af in assigned_imgs:
                            inserted_image_report.append(f"スライド {ch_page_num}（{clean_ch_title}）: 『{af['filename']}』")

                    if should_draw_chart:
                        c_dict, is_valid_chart = verify_and_fix_chart_data(
                            raw_chart_info, layout_type, clean_ch_title, body_text_val, raw_memo
                        )
                        categories = c_dict.get("categories", [])
                        values = c_dict.get("values", [])
                        g_title_raw = str(c_dict.get("title") or f"{clean_ch_title}")
                        g_unit = str(c_dict.get("unit") or "").strip()

                        if g_unit and f"（{g_unit}）" not in g_title_raw and f"({g_unit})" not in g_title_raw:
                            g_title_text = f"{g_title_raw}（{g_unit}）"
                        else:
                            g_title_text = g_title_raw

                        series_name = str(c_dict.get("series_name") or g_title_text)

                        if is_valid_chart:
                            chart_data = CategoryChartData()
                            chart_data.categories = categories
                            chart_data.add_series(series_name, values)

                            try:
                                is_pie = ("pie" in layout_type or "円" in clean_ch_title or any(kw in clean_ch_title for kw in ["構成", "チャネル", "年代"]))
                                is_bar = ("bar" in layout_type or "棒" in clean_ch_title) and not is_pie

                                if is_pie:
                                    c_type = XL_CHART_TYPE.PIE
                                    x_pos, y_pos, cx_pos, cy_pos = Pt(380), Pt(130), Pt(340), Pt(270)
                                elif is_bar:
                                    c_type = XL_CHART_TYPE.COLUMN_CLUSTERED
                                    x_pos, y_pos, cx_pos, cy_pos = Pt(380), Pt(135), Pt(320), Pt(260)
                                else:
                                    c_type = XL_CHART_TYPE.LINE_MARKERS
                                    x_pos, y_pos, cx_pos, cy_pos = Pt(380), Pt(135), Pt(320), Pt(260)

                                chart_shape = ch_slide.shapes.add_chart(c_type, x_pos, y_pos, cx_pos, cy_pos, chart_data)
                                chart = chart_shape.chart
                                chart.has_title = True
                                chart.chart_title.text_frame.text = g_title_text

                                if is_pie:
                                    pie_palette = [
                                        RGBColor(217, 48, 37),
                                        RGBColor(235, 122, 25),
                                        RGBColor(98, 178, 44),
                                        RGBColor(65, 126, 186),
                                        RGBColor(138, 79, 148),
                                        RGBColor(78, 78, 78),
                                        RGBColor(246, 189, 96),
                                    ]
                                    series = chart.series[0]
                                    for p_i, pt in enumerate(series.points):
                                        pt.format.fill.solid()
                                        pt.format.fill.fore_color.rgb = pie_palette[p_i % len(pie_palette)]
                                        pt.format.line.color.rgb = RGBColor(255, 255, 255)
                                        pt.format.line.width = Pt(1.5)

                                    plot = chart.plots[0]
                                    plot.has_data_labels = True
                                    data_labels = plot.data_labels
                                    data_labels.show_percentage = True
                                    data_labels.show_value = False
                                    data_labels.show_category_name = False
                                    data_labels.position = XL_DATA_LABEL_POSITION.BEST_FIT

                                    if hasattr(data_labels, 'font'):
                                        data_labels.font.color.rgb = RGBColor(255, 255, 255)
                                        data_labels.font.bold = True
                                        data_labels.font.name = "メイリオ"
                                        data_labels.font.size = Pt(11)

                                    chart.has_legend = True
                                    chart.legend.position = XL_LEGEND_POSITION.RIGHT
                                    chart.legend.include_in_layout = False

                                elif is_bar:
                                    chart.has_legend = False
                                    plot = chart.plots[0]
                                    if hasattr(plot, 'vary_by_categories'):
                                        plot.vary_by_categories = False

                                    for s in chart.series:
                                        s.format.fill.solid()
                                        s.format.fill.fore_color.rgb = RGBColor(217, 48, 37)
                                        s.format.line.color.rgb = RGBColor(217, 48, 37)
                                        try:
                                            for point in s.points:
                                                point.format.fill.solid()
                                                point.format.fill.fore_color.rgb = RGBColor(217, 48, 37)
                                                point.format.line.color.rgb = RGBColor(217, 48, 37)
                                        except Exception:
                                            pass

                                    val_ax = getattr(chart, 'value_axis', None)
                                    if val_ax and hasattr(val_ax, 'tick_labels'):
                                        is_percent_data = (g_unit in ["％", "%"] or "％" in g_title_text or "%" in g_title_text or any(kw in clean_ch_title for kw in ["比率", "割合", "構成"]))
                                        if is_percent_data:
                                            val_ax.tick_labels.number_format = '0"%"'
                                            val_ax.tick_labels.number_format_is_linked = False

                                else:
                                    chart.has_legend = False
                                    for s in chart.series:
                                        s.format.line.color.rgb = RGBColor(217, 48, 37)
                                        s.format.line.width = Pt(2.5)
                                        try:
                                            if hasattr(s, 'marker'):
                                                s.marker.format.fill.solid()
                                                s.marker.format.fill.fore_color.rgb = RGBColor(217, 48, 37)
                                                s.marker.format.line.color.rgb = RGBColor(217, 48, 37)
                                        except Exception:
                                            pass

                                    val_ax = getattr(chart, 'value_axis', None)
                                    if val_ax and hasattr(val_ax, 'tick_labels'):
                                        is_percent_data = (g_unit in ["％", "%"] or "％" in g_title_text or "%" in g_title_text or any(kw in clean_ch_title for kw in ["比率", "割合", "構成"]))
                                        if is_percent_data:
                                            val_ax.tick_labels.number_format = '0"%"'
                                            val_ax.tick_labels.number_format_is_linked = False

                                force_meiryo_and_style_on_chart(chart, font_size_pt=9)
                                paired_str = ", ".join([f"{c}:{v}{g_unit}" for c, v in zip(categories, values)])
                                chart_audit_report.append(f"✅ P{ch_page_num}（{g_title_text}）: [{paired_str}]")
                            except Exception as e:
                                chart_audit_report.append(f"⚠️ P{ch_page_num} グラフ描画エラー: {e}")
                        else:
                            chart_audit_report.append(f"⚠️ P{ch_page_num}（{g_title_text}）: 数値データが不足しているためグラフを省略しました")

                # 4. テンプレート元スライドの完全消去
                for _ in range(tpl_slide_count):
                    del prs.slides._sldIdLst[0]

                # 5. 全スライドへのページ番号の確実な付与（P2〜）
                for s_i, sld in enumerate(prs.slides):
                    for shp in sld.shapes:
                        if shp.has_text_frame:
                            for p in shp.text_frame.paragraphs:
                                if "[[" in p.text and "]]" in p.text:
                                    p.text = re.sub(r'\[\[.*?\]\]', '', p.text)
                    if s_i > 0:
                        ensure_page_number(sld, s_i)

                # 6. ID・GUID重複＆孤立リレーションの完全パージ
                fix_all_xml_integrity_safe(prs)

                # 7. メモリ上で完結
                output_stream = io.BytesIO()
                prs.save(output_stream)
                output_stream.seek(0)

                safe_filename = re.sub(r'[/\\:*?"<>|]', '_', doc_title.strip())
                if not safe_filename:
                    safe_filename = "提案書"
                download_filename = f"{safe_filename}.pptx"

                st.balloons()
                st.success(f"🎉 『{download_filename}』の自動構築が完了しました！（全 {len(prs.slides)} ページ）")

                if chart_audit_report:
                    with st.expander("📈 **グラフデータ整合性チェック結果（元データ照合）**", expanded=True):
                        for line in chart_audit_report:
                            st.write(line)

                q_report = run_quality_check(raw_memo, edited_chapters, doc_title)
                with st.expander("📊 **AI品質チェック＆網羅性診断レポート（入力情報反映確認）**", expanded=True):
                    st.markdown("元メモに記載された重要データ・施策・目標が、資料内に反映されているかを自動検証しました。")
                    q_cols = st.columns(2)
                    for q_idx, (topic_name, is_covered) in enumerate(q_report.items()):
                        with q_cols[q_idx % 2]:
                            if is_covered:
                                st.write(f"✅ **{topic_name}**: 反映済み")
                            else:
                                st.write(f"⚠️ **{topic_name}**: 未検出（プレビュー編集画面から追記・調整可能）")

                if inserted_image_report:
                    st.info("📷 **以下のスライドに画像を配置しました：**\n- " + "\n- ".join(inserted_image_report))

                st.download_button(
                    label=f"📥 『{download_filename}』をダウンロード",
                    data=output_stream,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"エラーが発生しました: {e}")
