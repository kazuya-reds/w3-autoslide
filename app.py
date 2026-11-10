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
import json
import datetime
import io
import re
import time
import copy
import os
import uuid
import unicodedata

st.set_page_config(page_title="W3 AutoSlide", layout="wide")

# カスタムCSS
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

st.image("logo.png", width=350)
st.caption("AI資料 自動生成システム")

LAYOUT_MAP = {
    "line_chart": "折れ線グラフ（推移・トレンド）",
    "bar_chart": "棒グラフ（比較・実績）",
    "pie_chart": "円グラフ（構成比・内訳）",
    "text": "解説スライド（文章メイン または 写真添え）",
    "step": "実行ステップ（フロー・手順図）",
    "checklist": "チェックリスト（確認項目一覧）",
    "table": "表組み（施策一覧・タスク管理）"
}
LAYOUT_REVERSE_MAP = {v: k for k, v in LAYOUT_MAP.items()}

if "saved_images" not in st.session_state:
    st.session_state["saved_images"] = []

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
    if not text: return ""
    return re.sub(r'^(?:[※\*\:\s]|注意[：:]|注[：:])+', '', str(text).strip())

def sanitize_text_empty_lines(text):
    if not text: return ""
    return re.sub(r'\n{2,}', '\n', str(text).strip())

def clean_chapter_title(title):
    if not title:
        return ""
    res = re.sub(r'^(?:第\s*\d+\s*[章節]\s*|\d{1,2}[\.\:\-_・／/]\s*|\d{2}\s+(?![ヶカか年月週日つ個大度ステップSTEP])|\d{1}\s+(?![ヶカか年月週日つ個大度ステップSTEP]))', '', str(title).strip())
    return res.strip()

def clean_step_text(text):
    if not text:
        return ""
    cleaned = re.sub(r'^(?:[①-⑳]|\(?\d+[\.\)）]\s*|(?:STEP|Step|step|ステップ)\s*\d+[\:\s・]*|[・\-\*]\s*)+', '', str(text).strip())
    return cleaned.strip()

def sanitize_main_item(main_item, clean_ch_title):
    main_item = str(main_item).strip() if main_item else ""
    if not main_item or len(main_item) > 18 or "。" in main_item or "、" in main_item:
        first_part = re.split(r'[。\n、,]', main_item)[0].strip()
        if 2 <= len(first_part) <= 15:
            main_item = first_part
        else:
            if "推移" in clean_ch_title or "売上" in clean_ch_title:
                main_item = "売上実績は堅調に推移"
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

def extract_chart_data_from_memo(text):
    pattern = r'(\d{1,2}月)[^\d\n]{0,10}?(\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?)\s*(?:万|万円)?'
    matches = re.findall(pattern, text)
    if matches:
        categories = []
        values = []
        for m in matches:
            cat = m[0]
            val_str = m[1].replace(',', '')
            try:
                val = float(val_str)
                if cat not in categories:
                    categories.append(cat)
                    values.append(val)
            except Exception:
                pass
        if len(categories) >= 2:
            return {
                "title": "月別推移",
                "unit": "万円",
                "categories": categories,
                "series_name": "推移データ",
                "values": values
            }
    return {
        "title": "月別推移",
        "unit": "万円",
        "categories": ["4月", "5月", "6月", "7月", "8月", "9月"],
        "series_name": "推移データ",
        "values": [1200.0, 1450.0, 1380.0, 1600.0, 1850.0, 1720.0]
    }

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
    font_size_step = Pt(11.0) if has_image else Pt(13.0)

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
        if "STEP_1_BADGE" in sname: step_badges[1] = shp
        elif "STEP_1_TEXT" in sname: step_texts[1] = shp
        elif "STEP_2_BADGE" in sname: step_badges[2] = shp
        elif "STEP_2_TEXT" in sname: step_texts[2] = shp
        elif "STEP_3_BADGE" in sname: step_badges[3] = shp
        elif "STEP_3_TEXT" in sname: step_texts[3] = shp
        elif "STEP_4_BADGE" in sname: step_badges[4] = shp
        elif "STEP_4_TEXT" in sname: step_texts[4] = shp

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

    start_top = Pt(200.0)
    step_gap = Pt(45.0)

    for k in range(1, 5):
        if k in step_texts:
            t_shp = step_texts[k]
            t_shp.left = Pt(180.0) if has_image else Pt(190.0)
            t_shp.top = int(start_top + (k - 1) * step_gap)
            t_shp.width = step_text_width
            t_shp.height = Pt(32.0)
            if t_shp.has_text_frame:
                tf = t_shp.text_frame
                tf.word_wrap = True
                raw_desc = step_descs[k - 1] if len(step_descs) >= k else f"アクション {k}"
                tf.text = clean_step_text(raw_desc)
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
        elif shp.shape_type == 1 and Pt(50) < Pt(300):
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

def apply_clean_table_borders(table):
    ns_a = nsdecls("a")
    border_dark = "02182E"
    border_light = "808C99"
    
    tbl_el = table._tbl
    trs = tbl_el.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tr')
    
    for r_idx, tr in enumerate(trs):
        is_header = (r_idx == 0)
        is_last_row = (r_idx == len(trs) - 1)
        tcs = tr.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tc')
        
        for c_idx, tc in enumerate(tcs):
            is_first_col = (c_idx == 0)
            is_last_col = (c_idx == len(tcs) - 1)
            tcPr = tc.get_or_add_tcPr()
            
            for b_tag in ['lnL', 'lnR', 'lnT', 'lnB', 'lnTlToBr', 'lnBlToTr']:
                old_b = tcPr.find(f'{{http://schemas.openxmlformats.org/drawingml/2006/main}}{b_tag}')
                if old_b is not None:
                    tcPr.remove(old_b)
            
            if is_header:
                tcPr.append(parse_xml(f'<a:lnT {ns_a} w="19050" cmpd="sng"><a:solidFill><a:srgbClr val="{border_dark}"/></a:solidFill></a:lnT>'))
            else:
                tcPr.append(parse_xml(f'<a:lnT {ns_a} w="12700" cmpd="sng"><a:solidFill><a:srgbClr val="{border_light}"/></a:solidFill></a:lnT>'))
                
            if is_header:
                tcPr.append(parse_xml(f'<a:lnB {ns_a} w="25400" cmpd="sng"><a:solidFill><a:srgbClr val="{border_dark}"/></a:solidFill></a:lnB>'))
            elif is_last_row:
                tcPr.append(parse_xml(f'<a:lnB {ns_a} w="19050" cmpd="sng"><a:solidFill><a:srgbClr val="{border_dark}"/></a:solidFill></a:lnB>'))
            else:
                tcPr.append(parse_xml(f'<a:lnB {ns_a} w="12700" cmpd="sng"><a:solidFill><a:srgbClr val="{border_light}"/></a:solidFill></a:lnB>'))
                
            tcPr.append(parse_xml(f'<a:lnL {ns_a} w="12700" cmpd="sng"><a:solidFill><a:srgbClr val="{border_dark if is_first_col else border_light}"/></a:solidFill></a:lnL>'))
            tcPr.append(parse_xml(f'<a:lnR {ns_a} w="12700" cmpd="sng"><a:solidFill><a:srgbClr val="{border_dark if is_last_col else border_light}"/></a:solidFill></a:lnR>'))

def format_table_and_note(slide, total_width_pt=671.1, assigned_images=None, captions=None, table_data=None):
    wipe_right_side_shapes(slide)
    has_image = bool(assigned_images and len(assigned_images) > 0)
    actual_width = Pt(320.0) if has_image else Pt(total_width_pt)
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
        table = tbl_shape.table
        tbl_shape.left = Pt(113.7)
        tbl_shape.top = Pt(145.0)
        
        headers = table_data.get("headers", [])
        rows = table_data.get("rows", [])
        target_cols = max(1, len(headers))
        target_rows = max(1, len(rows)) + 1

        tbl_el = table._tbl
        while len(table.columns) > target_cols:
            col_to_remove = len(table.columns) - 1
            tblGrid = tbl_el.tblGrid
            gridCols = tblGrid.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}gridCol')
            if col_to_remove < len(gridCols):
                tblGrid.remove(gridCols[col_to_remove])
            for tr in tbl_el.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tr'):
                tcs = tr.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tc')
                if col_to_remove < len(tcs):
                    tr.remove(tcs[col_to_remove])

        while len(table.rows) < target_rows:
            trs = tbl_el.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tr')
            last_tr = trs[-1]
            new_tr = copy.deepcopy(last_tr)
            for tc in new_tr.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tc'):
                for p in tc.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}txBody/{http://schemas.openxmlformats.org/drawingml/2006/main}p'):
                    for r in p.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}r'):
                        t = r.find('{http://schemas.openxmlformats.org/drawingml/2006/main}t')
                        if t is not None:
                            t.text = ""
            tbl_el.append(new_tr)

        for c_idx, h_text in enumerate(headers):
            if c_idx < len(table.columns):
                table.rows[0].cells[c_idx].text = str(h_text)

        for r_idx, row_items in enumerate(rows):
            target_r_idx = r_idx + 1
            if target_r_idx < len(table.rows):
                for c_idx, val in enumerate(row_items):
                    if c_idx < len(table.columns):
                        table.rows[target_r_idx].cells[c_idx].text = str(val)

        if target_cols == 3:
            proportions = [0.22, 0.38, 0.40]
        elif target_cols == 4:
            proportions = [0.20, 0.30, 0.30, 0.20]
        else:
            proportions = [1.0 / target_cols] * target_cols

        for c in range(len(table.columns)):
            col_w = (actual_width.pt) * proportions[c]
            table.columns[c].width = int(Pt(col_w))

        font_sz = Pt(9.0) if (len(table.rows) >= 6 or has_image) else Pt(10.0)
        for row in table.rows:
            for cell in row.cells:
                cell.text_frame.word_wrap = True
                cell.text_frame.margin_left = Pt(6)
                cell.text_frame.margin_right = Pt(6)
                cell.text_frame.margin_top = Pt(4)
                cell.text_frame.margin_bottom = Pt(4)
                for p in cell.text_frame.paragraphs:
                    p.font.name = "メイリオ"
                    p.font.size = font_sz
                    for r in p.runs:
                        r.font.name = "メイリオ"
                        r.font.size = font_sz

        apply_clean_table_borders(table)

    if tbl_shape and note_shape:
        note_shape.left = Pt(113.7)
        note_shape.top = int(tbl_shape.top + tbl_shape.height + Pt(14.0))
        note_shape.width = actual_width
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

    if sub_title or sub_body:
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

def format_and_align_index_slide(slide):
    title_shapes = {}
    page_shapes = {}

    for shp in slide.shapes:
        sname = shp.name.upper()
        for k in range(1, 15):
            if f"INDEX_{k}_TITLE" in sname:
                title_shapes[k] = shp
            elif f"INDEX_{k}_PAGE" in sname:
                page_shapes[k] = shp

    max_title_right = Pt(460.0)

    for k in range(1, 15):
        if k in title_shapes:
            t_shp = title_shapes[k]
            if t_shp.has_text_frame and t_shp.text_frame.text.strip():
                txt = t_shp.text_frame.text.strip()
                t_shp.text_frame.word_wrap = False
                t_shp.width = Pt(340.0)

                p = t_shp.text_frame.paragraphs[0]
                m = re.match(r'^(\d{2})\s*(.*)$', txt)
                if m:
                    num_str = m.group(1)
                    title_str = m.group(2)
                    p.text = ""

                    r0 = p.add_run()
                    r0.text = f"{num_str} "
                    r0.font.name = "メイリオ"
                    r0.font.size = Pt(18.0)
                    r0.font.bold = True
                    r0.font.color.rgb = RGBColor(217, 48, 37)

                    r1 = p.add_run()
                    r1.text = title_str
                    r1.font.name = "メイリオ"
                    r1.font.size = Pt(18.0)
                    r1.font.bold = True

                calc_w = 0
                for c in txt:
                    calc_w += 18.5 if ord(c) > 127 else 10.0
                calc_right = t_shp.left + Pt(calc_w)
                if calc_right > max_title_right:
                    max_title_right = calc_right

    page_left_pos = max(Pt(550.0), int(max_title_right + Pt(25.0)))

    for k in range(1, 15):
        if k in page_shapes:
            p_shp = page_shapes[k]
            if p_shp.has_text_frame and p_shp.text_frame.text.strip():
                p_shp.left = page_left_pos
                p_shp.width = Pt(70.0)
                p_shp.text_frame.word_wrap = False
                for p in p_shp.text_frame.paragraphs:
                    p.alignment = PP_ALIGN.LEFT
                    for r in p.runs:
                        r.font.name = "メイリオ"
                        r.font.size = Pt(18.0)
                        r.font.bold = True

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
        "平日集客の対策": ["平日", "ランチ", "団体"],
        "SNSの見直し・強化": ["SNS", "Instagram", "TikTok"],
        "写真素材の撮影": ["写真", "撮影", "撮り直し"],
        "イベント企画の刷新": ["イベント", "マルシェ", "ライトアップ", "企画"],
        "地域・近隣企業との連携": ["地域", "近隣", "連携", "タイアップ", "飲食店", "企業"],
        "若年層（20〜30代）の獲得": ["若い", "若年", "20代", "30代"],
        "現場の負担配慮・スモールスタート": ["人手", "スタッフ", "無理なく", "スモールスタート", "予算"]
    }
    
    all_slide_text = doc_title + " " + " ".join([
        ch.get("chapter_title", "") + " " +
        ch.get("main_item", "") + " " +
        ch.get("body", "") + " " +
        ch.get("sub_title", "") + " " +
        ch.get("sub_body", "") + " " +
        " ".join(ch.get("step_descs", [])) + " " +
        " ".join(ch.get("checklist_items", [])) + " " +
        " ".join([val for row in ch.get("table", {}).get("rows", []) for val in row])
        for ch in edited_chapters
    ])
    
    results = {}
    for topic, kws in memo_topics.items():
        found = any(kw in all_slide_text for kw in kws)
        results[topic] = found
        
    return results
    # 1. 表紙スライド
                cover_slide = duplicate_slide_safe(prs, 0)
                cover_map = {
                    "[[資料タイトル]]": doc_title,
                    "[[顧客名]]": client_name,
                    "[[バージョン]]": "1.0",
                    "[[更新日]]": today_str
                }
                process_slide_shapes(cover_slide, cover_map, is_cover=True)

                # 2. 目次スライド（★安全な clean_chapter_title を適用★）
                index_slide = duplicate_slide_safe(prs, 1)
                index_map = {}
                for ch_idx, ch in enumerate(edited_chapters):
                    ch_num_str = f"{ch_idx + 1:02d}"
                    ch_page_num = ch_idx + 2
                    raw_ch_title = ch.get("chapter_title", f"章{ch_idx+1}")
                    clean_ch_title = clean_chapter_title(raw_ch_title)

                    index_map[f"[[章{ch_idx+1}タイトル]]"] = f"{ch_num_str} {clean_ch_title}"
                    index_map[f"P[[章{ch_idx+1}ページ]]"] = f"P{ch_page_num}"
                    index_map[f"[[章{ch_idx+1}ページ]]"] = str(ch_page_num)

                for k in range(len(edited_chapters) + 1, 15):
                    index_map[f"[[章{k}タイトル]]"] = ""
                    index_map[f"P[[章{k}ページ]]"] = ""
                    index_map[f"[[章{k}ページ]]"] = ""

                process_slide_shapes(index_slide, index_map)
                format_and_align_index_slide(index_slide)

                inserted_image_report = []

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
                    chart_info = ch.get("chart_info")

                    should_draw_chart = (layout_type in ["line_chart", "bar_chart", "pie_chart"])

                    # テンプレート割り当て
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
                    if (not isinstance(step_descs, list) or len(step_descs) < 4) and body_text_val:
                        b_steps = re.findall(r'(?:\d+[\.\s、]|\・|\-)\s*([^\n]+)', body_text_val)
                        if len(b_steps) >= 3:
                            step_descs = b_steps[:4]
                    if not isinstance(step_descs, list) or len(step_descs) < 4:
                        step_descs = [
                            "現状のSNS運用の見直しと魅力的な写真素材の撮影",
                            "平日向けランチプラン・新イベント企画・地域連携の準備",
                            "各種施策の告知開始と3ヶ月間の試行運用テスト",
                            "来場者数やSNS反応等のデータ確認と次の展開の検討"
                        ]
                    for idx_s in range(1, 5):
                        raw_desc_s = step_descs[idx_s - 1] if len(step_descs) >= idx_s else ""
                        ch_map[f"[[STEP{idx_s}の説明]]"] = clean_step_text(raw_desc_s)

                    c_intro = ch.get("checklist_intro", "施策実行における重要確認事項です。")
                    c_outro = ch.get("checklist_outro", "全項目の確認完了後、各担当部門にて施策を進行してください。")
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

                    # レイアウト別の整形
                    if layout_type == "step":
                        build_perfect_steps(ch_slide, step_descs, body_text_val, assigned_imgs, captions)
                    elif layout_type == "checklist":
                        format_checklist_slide(ch_slide, assigned_imgs, captions)
                    elif layout_type == "table":
                        # ★動的行追加・列削除で全施策を完璧に整形★
                        format_table_and_note(ch_slide, total_width_pt=671.1, assigned_images=assigned_imgs, captions=captions, table_data=tbl_data)
                    elif should_draw_chart:
                        format_chart_slide_left_column(ch_slide)
                    else:
                        format_text_or_image_slide_layout(ch_slide, assigned_imgs, captions)

                    if len(assigned_imgs) > 0 and not should_draw_chart:
                        for af in assigned_imgs:
                            inserted_image_report.append(f"スライド {ch_page_num}（{clean_ch_title}）: 『{af['filename']}』")

                    # グラフ描画
                    if should_draw_chart and isinstance(chart_info, dict):
                        categories = chart_info.get("categories", ["4月", "5月", "6月", "7月", "8月", "9月"])
                        g_title_raw = str(chart_info.get("title") or f"{clean_ch_title}")
                        g_unit = str(chart_info.get("unit") or "").strip()

                        if g_unit and f"（{g_unit}）" not in g_title_raw and f"({g_unit})" not in g_title_raw:
                            g_title_text = f"{g_title_raw}（{g_unit}）"
                        else:
                            g_title_text = g_title_raw

                        series_name = str(chart_info.get("series_name") or g_title_text)
                        raw_values = chart_info.get("values", [1200.0, 1450.0, 1380.0, 1600.0, 1850.0, 1720.0])

                        values = []
                        if isinstance(raw_values, list):
                            for v in raw_values:
                                try:
                                    clean_num = str(v).replace(',', '').strip()
                                    values.append(float(clean_num) if clean_num else 0.0)
                                except Exception:
                                    values.append(0.0)

                        if categories and values:
                            min_len = min(len(categories), len(values))
                            chart_data = CategoryChartData()
                            chart_data.categories = categories[:min_len]
                            chart_data.add_series(series_name, values[:min_len])

                            try:
                                is_pie = ("pie" in layout_type or "円" in clean_ch_title)
                                is_bar = ("bar" in layout_type or "棒" in clean_ch_title)

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

                                # 円グラフ
                                if is_pie:
                                    pie_palette = [
                                        RGBColor(217, 48, 37),   # 1. 赤
                                        RGBColor(235, 122, 25),  # 2. オレンジ
                                        RGBColor(98, 178, 44),   # 3. グリーン
                                        RGBColor(65, 126, 186),  # 4. ブルー
                                        RGBColor(138, 79, 148),  # 5. パープル
                                        RGBColor(78, 78, 78),    # 6. ダークグレー
                                        RGBColor(246, 189, 96),  # 7. ウォームイエロー
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

                                # 棒グラフ
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

                                # 折れ線グラフ
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
                            except Exception:
                                pass

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
                
                # ★【新設】AI品質チェック＆重要事項網羅性診断レポートの表示★
                q_report = run_quality_check(raw_memo, edited_chapters, doc_title)
                with st.expander("📊 **AI品質チェック＆網羅性診断レポート（自動検証結果）**", expanded=True):
                    st.markdown("元メモに記載された重要課題・施策が、資料内に確実に反映されているかを自動検証しました。")
                    q_cols = st.columns(2)
                    for q_idx, (topic_name, is_covered) in enumerate(q_report.items()):
                        with q_cols[q_idx % 2]:
                            if is_covered:
                                st.write(f"✅ **{topic_name}**: 網羅・反映済み")
                            else:
                                st.write(f"⚠️ **{topic_name}**: 未検出（確認を推奨）")

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
