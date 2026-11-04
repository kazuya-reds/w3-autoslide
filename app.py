import streamlit as st
from google import genai
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.chart import XL_CHART_TYPE
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
    "image": "画像付き解説（画像＋テキスト）",
    "text": "標準テキスト（文章・概要）",
    "step": "実行ステップ（フロー・手順図）",
    "checklist": "チェックリスト（確認項目一覧）",
    "table": "表組み（タスク・管理一覧）"
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

def sanitize_main_item(main_item, clean_ch_title):
    main_item = str(main_item).strip() if main_item else ""
    if not main_item or len(main_item) > 18 or "。" in main_item or "、" in main_item:
        first_part = re.split(r'[。\n、,]', main_item)[0].strip()
        if 2 <= len(first_part) <= 15:
            main_item = first_part
        else:
            if "推移" in clean_ch_title or "売上" in clean_ch_title or "販売" in clean_ch_title:
                main_item = "売上実績は堅調に推移"
            elif "要因" in clean_ch_title or "分析" in clean_ch_title:
                main_item = "実績を牽引した主な要因"
            elif "アクション" in clean_ch_title or "手順" in clean_ch_title or "戦略" in clean_ch_title or "施策" in clean_ch_title:
                main_item = "今後の重要アクション"
            elif "チェック" in clean_ch_title or "確認" in clean_ch_title or "課題" in clean_ch_title:
                main_item = "現状の主要課題と確認事項"
            elif "表" in clean_ch_title or "タスク" in clean_ch_title or "管理" in clean_ch_title:
                main_item = "実行タスクの進捗一覧"
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
                "title": "月別売上推移（万円）",
                "categories": categories,
                "series_name": "売上高（万円）",
                "values": values
            }
    return {
        "title": "月別売上推移（万円）",
        "categories": ["4月", "5月", "6月", "7月", "8月", "9月"],
        "series_name": "売上高（万円）",
        "values": [1200.0, 1450.0, 1380.0, 1600.0, 1850.0, 1720.0]
    }

# ★第1防衛線：複製時点でGUIDとIDを完全に新品化する安全複製エンジン★
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
        new_el = copy.deepcopy(shape._element)
        # 複製直後に creationId を書き換えて重複を根絶
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
        is_name = any(kw in shp.name.lower() for kw in ["スライド番号", "page_number", "pagenumber", "slide_number"])
        if is_ph or is_name:
            found_shp = shp
            break

    if found_shp and found_shp.has_text_frame:
        tf = found_shp.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.text = str(page_num)
        p.alignment = PP_ALIGN.RIGHT
        p.font.name = "メイリオ"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
    else:
        txBox = slide.shapes.add_textbox(Pt(720), Pt(542), Pt(80), Pt(28))
        tf = txBox.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.text = str(page_num)
        p.alignment = PP_ALIGN.RIGHT
        p.font.name = "メイリオ"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)

# ★第2・第3防衛線：OpenXML整合性完全修復エンジン★
def fix_all_xml_integrity_safe(prs):
    global_id = 100
    for sld in prs.slides:
        tree = sld.shapes._spTree
        for node in tree.iter():
            # 1. cNvPr id を完全一意な連番に更新
            if node.tag.endswith('cNvPr'):
                global_id += 1
                node.attrib['id'] = str(global_id)
            # 2. creationId の重複を完全に解消
            elif node.tag.endswith('creationId'):
                node.attrib['id'] = f"{{{str(uuid.uuid4()).upper()}}}"

    # 3. 削除されたスライドの孤立リレーション（幽霊リンク）を完全消去
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

    # 4. 各スライド内の孤立リレーションも完全消去
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
            p.font.name = "メイリオ"
            p.font.size = Pt(11)
            for run in p.runs:
                run.font.name = "メイリオ"
                run.font.size = Pt(11)

    for axis in [getattr(chart, 'category_axis', None), getattr(chart, 'value_axis', None)]:
        if axis and hasattr(axis, 'tick_labels'):
            try:
                axis.tick_labels.font.name = "メイリオ"
                axis.tick_labels.font.size = Pt(font_size_pt)
            except Exception:
                pass

def wipe_right_side_shapes(slide):
    shapes_to_del = []
    for shp in slide.shapes:
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

def build_perfect_steps(slide, step_descs, body_text=""):
    wipe_right_side_shapes(slide)
    
    shapes_to_del = []
    main_body_shp = None
    for shp in slide.shapes:
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

    if main_body_shp and main_body_shp.has_text_frame:
        main_body_shp.left = Pt(113.7)
        main_body_shp.top = Pt(146.0)
        main_body_shp.width = Pt(660.0)
        main_body_shp.height = Pt(36.0)
        tf_b = main_body_shp.text_frame
        tf_b.word_wrap = True
        tf_b.text = clean_body
        for p in tf_b.paragraphs:
            p.font.name = "メイリオ"
            p.font.size = Pt(11.5)
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
            t_shp.left = Pt(190.0)
            t_shp.top = int(start_top + (k - 1) * step_gap)
            t_shp.width = Pt(580.0)
            t_shp.height = Pt(32.0)
            if t_shp.has_text_frame:
                tf = t_shp.text_frame
                tf.word_wrap = True
                tf.text = step_descs[k - 1] if len(step_descs) >= k else f"アクション {k}"
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                for p in tf.paragraphs:
                    p.font.size = Pt(13.0)
                    p.font.name = "メイリオ"
                    for r in p.runs:
                        r.font.size = Pt(13.0)
                        r.font.name = "メイリオ"

        if k in step_badges:
            b_shp = step_badges[k]
            b_shp.left = Pt(113.7)
            b_shp.height = Pt(22.0)
            b_shp.width = Pt(65.0)
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
                    p.font.size = Pt(9.5)
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(255, 255, 255)

def format_checklist_slide(slide):
    wipe_right_side_shapes(slide)
    main_title = None
    main_line = None
    intro_shape = None
    box_shape = None
    items_shape = None
    outro_shape = None

    shapes_to_del = []
    for shp in slide.shapes:
        sname = shp.name.upper()
        if "MAIN_SECTION_TITLE" in sname:
            main_title = shp
        elif "SECTION_LINE" in sname:
            main_line = shp
        elif "CHECKLIST_INTRO" in sname:
            intro_shape = shp
        elif shp.shape_type == 1 and shp.width > Pt(500) and shp.height > Pt(100):
            box_shape = shp
        elif "CHECKLIST_ITEMS" in sname or (shp.has_text_frame and "□" in shp.text_frame.text):
            items_shape = shp
        elif "CHECKLIST_OUTRO" in sname:
            outro_shape = shp
        elif any(kw in sname for kw in ["CAUTION", "SUPPLEMENT", "注意", "補足", "CAPTION"]):
            shapes_to_del.append(shp)
        elif shp.shape_type == 1 and Pt(50) < shp.width < Pt(500):
            shapes_to_del.append(shp)

    for shp in shapes_to_del:
        try:
            sp = shp._element
            sp.getparent().remove(sp)
        except Exception:
            pass

    if main_title:
        main_title.left = Pt(124.6)
        main_title.top = Pt(115.0)
        main_title.width = Pt(650.0)
        if main_line:
            main_line.left = Pt(113.7)
            main_line.top = Pt(120.5)

    if intro_shape:
        intro_shape.left = Pt(113.7)
        intro_shape.top = Pt(150.0)
        intro_shape.width = Pt(660.0)

    if box_shape:
        box_shape.left = Pt(113.7)
        box_shape.top = Pt(185.0)
        box_shape.width = Pt(660.0)
        box_shape.height = Pt(200.0)

    if items_shape and items_shape.has_text_frame and box_shape:
        items_shape.left = Pt(140.0)
        items_shape.width = Pt(610.0)
        tf = items_shape.text_frame
        valid_paras = [p for p in tf.paragraphs if p.text.strip()]
        num_items = max(1, len(valid_paras))
        
        for p in tf.paragraphs:
            p.font.name = "メイリオ"
            p.font.size = Pt(12.0)
            p.line_spacing = 1.3
            p.space_after = Pt(8.0)
            for r in p.runs:
                r.font.name = "メイリオ"
                r.font.size = Pt(12.0)

        effective_items_h = Pt(num_items * 28.0)
        items_shape.height = int(effective_items_h)
        items_shape.top = int(box_shape.top + (box_shape.height - effective_items_h) / 2)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    if outro_shape and box_shape:
        outro_shape.left = Pt(113.7)
        outro_shape.top = int(box_shape.top + box_shape.height + Pt(14.0))
        outro_shape.width = Pt(660.0)

def format_table_and_note(slide, total_width_pt=671.1):
    wipe_right_side_shapes(slide)
    tbl_shape = None
    note_shape = None

    shapes_to_del = []
    for shp in slide.shapes:
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

    if tbl_shape:
        tbl_shape.left = Pt(113.7)
        tbl_shape.top = Pt(150.0)
        for row in tbl_shape.table.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    p.font.name = "メイリオ"
                    p.font.size = Pt(11.0)
                    for r in p.runs:
                        r.font.name = "メイリオ"
                        r.font.size = Pt(11.0)

    if tbl_shape and note_shape:
        note_shape.left = Pt(113.7)
        note_shape.top = int(tbl_shape.top + tbl_shape.height + Pt(16.0))
        if note_shape.has_text_frame:
            for p in note_shape.text_frame.paragraphs:
                p.font.name = "メイリオ"
                p.font.size = Pt(9.5)
                p.font.color.rgb = RGBColor(120, 120, 120)

def format_chart_slide_left_column(slide):
    main_line = None
    main_title = None
    main_body = None
    sub_line = None
    sub_title = None
    sub_body = None

    for shp in slide.shapes:
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

def format_image_slide_layout(slide, assigned_images, captions=None):
    main_title = None
    main_line = None
    main_body = None
    sub_title = None
    sub_line = None
    sub_body = None

    for shp in slide.shapes:
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
        t_h = Pt(26.0)
        main_title.height = int(t_h)
        if main_line:
            main_line.left = Pt(113.7)
            main_line.top = int(curr_y + Pt(4.5))
        curr_y += t_h + Pt(12.0)

    if main_body and main_body.has_text_frame and main_body.text_frame.text.strip():
        main_body.left = Pt(113.7)
        main_body.top = int(curr_y)
        main_body.width = Pt(245.0)
        for p in main_body.text_frame.paragraphs:
            p.font.name = "メイリオ"
            p.font.size = Pt(11.0)
            p.line_spacing = 1.25
        b_lines = len(main_body.text_frame.text.strip()) // 16 + 1
        b_h = Pt(b_lines * 18.0)
        curr_y += b_h + Pt(14.0)

    has_sub = (sub_title and sub_title.has_text_frame and sub_title.text_frame.text.strip()) or \
              (sub_body and sub_body.has_text_frame and sub_body.text_frame.text.strip())

    if has_sub:
        if sub_title and sub_title.has_text_frame and sub_title.text_frame.text.strip():
            sub_title.left = Pt(124.6)
            sub_title.top = int(curr_y)
            sub_title.width = Pt(230.0)
            if sub_line:
                sub_line.left = Pt(113.7)
                sub_line.top = int(curr_y + Pt(3.5))
            curr_y += Pt(22.0)
        if sub_body and sub_body.has_text_frame and sub_body.text_frame.text.strip():
            sub_body.left = Pt(113.7)
            sub_body.top = int(curr_y)
            sub_body.width = Pt(245.0)
            for p in sub_body.text_frame.paragraphs:
                p.font.name = "メイリオ"
                p.font.size = Pt(10.5)
                p.line_spacing = 1.2
    else:
        for del_shp in [sub_line, sub_title, sub_body]:
            if del_shp:
                try:
                    sp = del_shp._element
                    sp.getparent().remove(sp)
                except Exception:
                    pass

    if assigned_images and len(assigned_images) > 0:
        img_bytes = assigned_images[0]["bytes"]
        pil_img = Image.open(io.BytesIO(img_bytes))
        orig_w, orig_h = pil_img.size

        max_w = Pt(360)
        max_h = Pt(300)
        scale = min(max_w / orig_w, max_h / orig_h)
        new_w = int(orig_w * scale)
        new_h = int(orig_h * scale)
        new_l = int(Pt(390) + (max_w - new_w) / 2)
        new_t = int(Pt(130) + (max_h - new_h) / 2)

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
    clean_title = re.sub(r'^\d+[\.\s_]*', '', str(ch_title)).strip() if ch_title else ""
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
            sname_u = shape.name.upper()
            raw_t = shape.text_frame.text if shape.has_text_frame else ""
            if any(kw in sname_u for kw in ["CAUTION", "SUPPLEMENT", "補足", "注意"]):
                shapes_to_remove.append(shape)
            elif any(kw in raw_t for kw in ["[[注意文]]", "[[補足文]]", "注意", "補足", "※", "注"]):
                shapes_to_remove.append(shape)

    for shape in slide.shapes:
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
        st.rerun()

col1, col2 = st.columns(2)
with col1:
    client_name = st.text_input("提案先（顧客名）", "例）役員会 各位")
    doc_title = st.text_input("資料タイトル", "例）売上推移と今後の成長戦略")
with col2:
    sender_name = st.text_input("作成者・部署など", "例）営業部")
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

raw_memo = st.text_area("資料の内容メモ", 
"""直近4ヶ月の売上推移について報告します。
7月は100万、8月は150万、9月は130万、10月は200万と堅調に推移しています。

売上推移の要因：
・8月および10月の増収は新規施策の奏功
・9月の一時的な落ち込みは季節要因と期ずれ

今後のアクション：
1. 顧客購買データの詳細分析
2. 高単価ターゲット層へのアプローチ強化
3. 新プロモーションの実施
4. 月商200万円の継続的維持

確認事項チェックリスト：
・顧客購買データのクレンジングと抽出完了
・高単価ターゲット層のセグメント確定
・プロモーション施策のKPIおよび予算決裁
・月次売上進捗レビュー体制の構築

実行タスク管理表：
施策項目, 担当部門, 完了期日, 進捗状況
データ分析基盤の構築, DX推進部, 11月15日, 順調
ターゲット層への提案骨子確定, 営業推進室, 11月30日, 着手済
新プロモーション展開, マーケティング部, 12月20日, 企画中""", height=240)


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
                    あなたはプレゼン資料作成のプロコンサルタントです。
                    以下の入力を直接分析し、メモの情報を一切落とさずにプレゼンスライド用のJSONを作成してください。

                    【入力データ】
                    {json.dumps(prompt_data, ensure_ascii=False, indent=2)}

                    【レイアウト選定ルール】
                    1. メモの主要項目を整理し、プレゼンに最適な章（chapters）にまとめてください。
                    2. 各章のコンテンツに合わせて、最適な "layout_type" を選定してください:
                       - "line_chart": 売上推移などの数値変化がある場合（※画像なし）
                       - "image": 画像を大きく表示して説明したいスライド（※画像＋テキスト解説）
                       - "step": 導入手順・運用フロー・実行ステップ（※画像なし・全幅STEP図）
                       - "checklist": 確認事項・必要書類・チェックリスト（※画像なし・全幅チェック欄）
                       - "table": タスク管理表、仕様一覧などの表組み（※画像なし・全幅表）
                       - "text": 概要・特徴・方針など
                    3. ★画像割り当ての絶対ルール★:
                       - アップロードされた画像が {len(image_contexts)} 枚あります。
                       - "step", "checklist", "table", "line_chart" には画像を割り当てないでください。
                       - アップロードされた全画像（インデックス 0〜{len(image_contexts)-1}）を1枚残らず使い切るため、画像用の章（"layout_type": "image"）を画像の枚数分だけ【新規に作成】し、それぞれの画像インデックスを assigned_image_indices に割り当ててください。
                    4. "step" の場合:
                       - "step_descs": メモから抽出した具体的な4つのアクションステップ文字列配列（例: ["1の施策", "2の施策", "3の施策", "4の施策"]）。
                    5. "checklist" の場合:
                       - "checklist_intro": 導入文
                       - "checklist_items": メモから抽出した最大6つのチェック項目文字列配列
                       - "checklist_outro": 完了後の案内文
                    6. "table" の場合:
                       - "table": {{
                           "headers": ["列1", "列2", "列3", "列4"],
                           "rows": [["データ1-1", "データ1-2", "データ1-3", "データ1-4"]],
                           "note": "※注記（任意）"
                         }}
                    7. "main_item" について:
                       - 15文字以内の短いキャッチコピーを指定してください。
                    8. "chart_info" について:
                       - 売上や推移データがある場合は "chart_info"（title, categories, series_name, values）を必ず格納してください。values は数値（浮動小数点または整数）の配列にしてください。

                    【出力形式】
                    純粋なJSONのみを出力してください。
                    {{
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
                          "step_descs": ["ステップ1", "ステップ2", "ステップ3", "ステップ4"],
                          "checklist_intro": "確認事項の導入文",
                          "checklist_items": ["項目1", "項目2", "項目3", "項目4"],
                          "checklist_outro": "確認完了後の案内文",
                          "table": {{
                            "headers": ["施策項目", "担当部門", "完了期日", "状況"],
                            "rows": [["タスク1", "担当A", "11/15", "着手済"]],
                            "note": "※注記文"
                          }},
                          "chart_info": {{
                            "title": "月別売上推移（万円）",
                            "categories": ["4月", "5月", "6月", "7月", "8月", "9月"],
                            "series_name": "売上高（万円）",
                            "values": [1200.0, 1450.0, 1380.0, 1600.0, 1850.0, 1720.0]
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
                    st.success("✅ AIの構成解析が完了しました！下部の編集フォームで内容を確認・微調整できます。")
                except Exception as e:
                    st.error(f"解析中にエラーが発生しました: {e}")

# ==========================================
# ステップ2: WEB編集フォーム ＆ 最終生成
# ==========================================
if st.session_state.get("analyzed", False) and "parsed_data" in st.session_state:
    st.subheader("✏️ AI解析結果のプレビュー ＆ WEB編集フォーム")
    st.caption("各スライドの内容を自由に変更できます。")

    parsed_chapters = st.session_state["parsed_data"].get("chapters", [])
    edited_chapters = []

    for idx, ch in enumerate(parsed_chapters):
        with st.expander(f"📄 スライド {idx + 1}: {ch.get('chapter_title', '')}", expanded=(idx <= 1)):
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
                
                selected_label = st.selectbox(f"レイアウトタイプ", labels_list, index=l_idx, key=f"layout_{idx}")
                e_layout = LAYOUT_REVERSE_MAP[selected_label]
                
                e_sub_title = st.text_input(f"補助項目 タイトル", value=ch.get("sub_title", ""), key=f"sub_t_{idx}")
                e_sub_body = st.text_area(f"補助項目 本文", value=sanitize_text_empty_lines(ch.get("sub_body", "")), height=68, key=f"sub_b_{idx}")

            e_assigned_indices = []
            e_captions = []
            if e_layout in ["image", "text"] and active_images:
                st.markdown("**📷 画像の挿入設定（任意）**")
                img_options = ["（挿入しない）"] + [f"画像{i+1}: {img['filename']}" for i, img in enumerate(active_images)]
                
                ai_assigned = ch.get("assigned_image_indices", [])
                default_sel_idx = 0
                if ai_assigned and len(ai_assigned) > 0 and ai_assigned[0] < len(active_images):
                    default_sel_idx = ai_assigned[0] + 1
                
                sel_img = st.selectbox("このスライドに挿入する画像", img_options, index=default_sel_idx, key=f"img_sel_{idx}")
                if sel_img != "（挿入しない）":
                    chosen_idx = img_options.index(sel_img) - 1
                    e_assigned_indices = [chosen_idx]
                    old_caps = ch.get("image_captions", [])
                    default_cap = old_caps[0] if (old_caps and isinstance(old_caps, list) and len(old_caps) > 0) else ""
                    user_cap = st.text_input("画像のキャプション（任意）", value=default_cap, key=f"cap_{idx}")
                    e_captions = [user_cap]
                    st.success(f"✅ このスライドの右側に 『{active_images[chosen_idx]['filename']}』 を挿入します")
            else:
                e_assigned_indices = []
                e_captions = []

            # 1. STEP詳細フォーム
            if e_layout == "step":
                st.markdown("**ステップ詳細（最大4ステップ）**")
                s_descs = ch.get("step_descs", [])
                if (not isinstance(s_descs, list) or len(s_descs) < 4) and e_body:
                    b_steps = re.findall(r'(?:\d+[\.\s、]|\・|\-)\s*([^\n]+)', e_body)
                    if len(b_steps) >= 3:
                        s_descs = b_steps[:4]
                if not isinstance(s_descs, list) or len(s_descs) < 4:
                    s_descs = [
                        "顧客購買データの詳細分析",
                        "高単価ターゲット層へのアプローチ強化",
                        "新プロモーションの実施",
                        "月商200万円の継続的維持"
                    ]
                e_s_descs = []
                s_cols = st.columns(2)
                for s_i in range(4):
                    default_s = s_descs[s_i] if len(s_descs) > s_i else f"アクション {s_i+1}"
                    with s_cols[s_i % 2]:
                        val_s = st.text_input(f"STEP {s_i+1}", value=default_s, key=f"step_{idx}_{s_i}")
                        e_s_descs.append(val_s if val_s else f"ステップ {s_i+1}")
            else:
                e_s_descs = ch.get("step_descs", [])

            # 2. チェックリスト詳細フォーム
            if e_layout == "checklist":
                st.markdown("**✅ チェックリスト設定（最大6項目）**")
                c_intro_val = st.text_input("チェックリスト導入文", value=ch.get("checklist_intro", "施策実行における重要確認事項です。"), key=f"c_intro_{idx}")
                c_items_def = ch.get("checklist_items", [])
                if not isinstance(c_items_def, list) or not c_items_def:
                    c_items_def = [
                        "店舗ごとの販売実績にばらつきがある",
                        "若年層の新規顧客獲得が不足している",
                        "中古車在庫の回転率に改善の余地がある",
                        "Webからの問い合わせに対する成約率が低い"
                    ]
                e_c_items = []
                ci_cols = st.columns(2)
                for ci in range(6):
                    def_item = c_items_def[ci] if len(c_items_def) > ci else ""
                    with ci_cols[ci % 2]:
                        val_item = st.text_input(f"チェック項目 {ci+1}", value=def_item, key=f"ci_{idx}_{ci}")
                        if val_item.strip():
                            e_c_items.append(val_item.strip())
                c_outro_val = st.text_input("確認後の案内文", value=ch.get("checklist_outro", "全項目の確認完了後、各担当部門にて施策を進行してください。"), key=f"c_outro_{idx}")
            else:
                c_intro_val = ch.get("checklist_intro", "")
                e_c_items = ch.get("checklist_items", [])
                c_outro_val = ch.get("checklist_outro", "")

            # 3. テーブル詳細フォーム（NoneTypeガード済み）
            if e_layout == "table":
                st.markdown("**📋 表組み（テーブル）データ設定**")
                tbl_raw = ch.get("table")
                if not isinstance(tbl_raw, dict):
                    tbl_raw = {}
                h_list = tbl_raw.get("headers")
                if not isinstance(h_list, list) or not h_list:
                    h_list = ["施策項目", "担当部門", "完了期日", "進捗状況"]
                r_list = tbl_raw.get("rows")
                if not isinstance(r_list, list) or not r_list:
                    r_list = [
                        ["データ分析基盤の構築", "DX推進部", "11月15日", "順調"],
                        ["ターゲット層への提案骨子確定", "営業推進室", "11月30日", "着手済"],
                        ["新プロモーション展開", "マーケティング部", "12月20日", "企画中"]
                    ]
                t_cols_input = st.text_input("列見出し（カンマ区切り・2〜4列）", value=", ".join(map(str, h_list)), key=f"th_{idx}")
                e_headers = [c.strip() for c in t_cols_input.split(",") if c.strip()]
                st.caption("各行のデータ（カンマ区切り）")
                e_rows = []
                for ri in range(3):
                    default_row_str = ", ".join(map(str, r_list[ri])) if len(r_list) > ri and isinstance(r_list[ri], list) else ""
                    row_str = st.text_input(f"{ri+1}行目", value=default_row_str, key=f"tr_{idx}_{ri}")
                    if row_str.strip():
                        e_rows.append([c.strip() for c in row_str.split(",")])
                t_note = st.text_input("表の注記（任意）", value=tbl_raw.get("note", "※進捗状況は週次定例にて更新します。"), key=f"tn_{idx}")
                e_table = {"headers": e_headers, "rows": e_rows, "note": t_note}
            else:
                e_table = ch.get("table") if isinstance(ch.get("table"), dict) else {}

            # 4. グラフ設定フォーム
            c_info = ch.get("chart_info")
            if not isinstance(c_info, dict) or not c_info.get("values"):
                c_info = extract_chart_data_from_memo(raw_memo)

            if e_layout in ["line_chart", "bar_chart"]:
                st.markdown("**📊 グラフデータ設定**")
                g_cols = st.columns(3)
                with g_cols[0]:
                    g_title = st.text_input("グラフタイトル", value=c_info.get("title", "月別売上推移（万円）"), key=f"gt_{idx}")
                with g_cols[1]:
                    g_cats_str = st.text_input("カテゴリ（カンマ区切り）", value=",".join(map(str, c_info.get("categories", ["4月","5月","6月","7月","8月","9月"]))), key=f"gc_{idx}")
                with g_cols[2]:
                    clean_vals = [str(v).replace(',', '') for v in c_info.get("values", [1200, 1450, 1380, 1600, 1850, 1720])]
                    g_vals_str = st.text_input("数値データ（カンマ区切り）", value=",".join(clean_vals), key=f"gv_{idx}")

                cats_list = [c.strip() for c in g_cats_str.split(",") if c.strip()]
                vals_list = []
                for v in g_vals_str.split(","):
                    try:
                        vals_list.append(float(re.sub(r'[^\d\.]', '', v.replace(',', ''))))
                    except Exception:
                        pass
                
                edited_c_info = {
                    "title": g_title,
                    "categories": cats_list,
                    "series_name": c_info.get("series_name", "売上高（万円）"),
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

    # 未割り当て画像を100%全件スライド化
    if active_images:
        assigned_indices = set()
        for ch in edited_chapters:
            for a_idx in ch.get("assigned_image_indices", []):
                if isinstance(a_idx, int) and 0 <= a_idx < len(active_images):
                    assigned_indices.add(a_idx)
        
        unassigned_indices = [i for i in range(len(active_images)) if i not in assigned_indices]
        for un_i in unassigned_indices:
            img_data = active_images[un_i]
            img_name = img_data.get("filename", f"画像{un_i+1}")
            img_title = os.path.splitext(img_name)[0]
            edited_chapters.append({
                "chapter_title": f"{img_title}の概要",
                "layout_type": "image",
                "main_item": f"{img_title}の活用と展開",
                "body": f"本画像（{img_name}）に関する詳細説明および関連施策の概要です。",
                "sub_title": "",
                "sub_body": "",
                "notice": "",
                "assigned_image_indices": [un_i],
                "image_captions": [img_title],
                "step_descs": [],
                "checklist_intro": "",
                "checklist_items": [],
                "checklist_outro": "",
                "table": {},
                "chart_info": {}
            })

    st.markdown("---")
    
    if st.button("✨ この内容で PowerPoint 資料を自動生成する ✨", key="generate_pptx_btn", type="primary", use_container_width=True):
        status_box = st.empty()
        with st.spinner("PowerPoint資料を精密構築中..."):
            try:
                current_dir = os.path.dirname(os.path.abspath(__file__))
                template_path = os.path.join(current_dir, '資料作成テンプレート_A4横.pptx')
                if not os.path.exists(template_path):
                    for f in os.listdir(current_dir):
                        if f.endswith('.pptx'):
                            template_path = os.path.join(current_dir, f)
                            break

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
                index_map = {}
                for ch_idx, ch in enumerate(edited_chapters):
                    ch_num_str = f"{ch_idx + 1:02d}"
                    ch_page_num = ch_idx + 2
                    ch_title = ch.get("chapter_title", f"章{ch_idx+1}")
                    clean_ch_title = re.sub(r'^\d+[\.\s_]*', '', str(ch_title)).strip()

                    index_map[f"[[章{ch_idx+1}タイトル]]"] = f"{ch_num_str} {clean_ch_title}"
                    index_map[f"P[[章{ch_idx+1}ページ]]"] = f"P{ch_page_num}"
                    index_map[f"[[章{ch_idx+1}ページ]]"] = str(ch_page_num)

                for k in range(len(edited_chapters) + 1, 15):
                    index_map[f"[[章{k}タイトル]]"] = ""
                    index_map[f"P[[章{k}ページ]]"] = ""
                    index_map[f"[[章{k}ページ]]"] = ""

                process_slide_shapes(index_slide, index_map)
                format_and_align_index_slide(index_slide)

                chart_created = False
                inserted_image_report = []

                # 3. 各章スライド構築
                for ch_idx, ch in enumerate(edited_chapters):
                    ch_num_str = f"{ch_idx + 1:02d}"
                    ch_page_num = ch_idx + 2
                    ch_title = ch.get("chapter_title", f"章{ch_idx+1}")
                    clean_ch_title = re.sub(r'^\d+[\.\s_]*', '', str(ch_title)).strip()

                    layout_type = str(ch.get("layout_type", "text")).lower()
                    img_indices = ch.get("assigned_image_indices", [])
                    assigned_imgs = [active_images[i] for i in img_indices if i < len(active_images)]
                    captions = ch.get("image_captions", [])

                    chart_info = ch.get("chart_info")
                    is_chart_keyword = any(kw in clean_ch_title for kw in ["推移", "売上", "実績", "販売", "比較", "グラフ", "データ"])

                    should_draw_chart = (not chart_created) and (
                        "line" in layout_type or "bar" in layout_type or
                        (is_chart_keyword and ch_idx == 0)
                    )

                    if should_draw_chart:
                        target_template_idx = 2
                    elif layout_type == "checklist":
                        target_template_idx = 9 if tpl_slide_count > 9 else 2
                    elif layout_type == "step":
                        target_template_idx = 5 if tpl_slide_count > 5 else 2
                    elif layout_type == "table":
                        target_template_idx = 8 if tpl_slide_count > 8 else 2
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
                            "月次販売実績の可視化と分析",
                            "店舗別の成功事例共有",
                            "顧客属性に応じた販促施策の実施",
                            "営業スタッフ向け研修の定期開催"
                        ]
                    for idx_s in range(1, 5):
                        desc_s = step_descs[idx_s - 1] if len(step_descs) >= idx_s else ""
                        ch_map[f"[[STEP{idx_s}の説明]]"] = desc_s

                    c_intro = ch.get("checklist_intro", "施策実行における重要確認事項です。")
                    c_outro = ch.get("checklist_outro", "全項目の確認完了後、各担当部門にて施策を進行してください。")
                    c_items = ch.get("checklist_items", [])
                    ch_map["[[チェックリスト導入文]]"] = c_intro
                    ch_map["[[チェック後の案内文]]"] = c_outro
                    for k in range(1, 7):
                        ch_map[f"[[チェック項目{k}]]"] = c_items[k - 1] if (isinstance(c_items, list) and len(c_items) >= k) else ""

                    tbl_data = ch.get("table")
                    if not isinstance(tbl_data, dict):
                        tbl_data = {}
                    headers = tbl_data.get("headers")
                    if not isinstance(headers, list) or not headers:
                        headers = ["施策項目", "担当部門", "完了期日", "進捗状況"]
                    rows = tbl_data.get("rows")
                    if not isinstance(rows, list) or not rows:
                        rows = [
                            ["データ分析基盤の構築", "DX推進部", "11月15日", "順調"],
                            ["ターゲット層への提案骨子確定", "営業推進室", "11月30日", "着手済"],
                            ["新プロモーション展開", "マーケティング部", "12月20日", "企画中"]
                        ]
                    ch_map["[[表の注記]]"] = tbl_data.get("note", ch.get("notice", ""))
                    for c_idx in range(1, 5):
                        ch_map[f"[[列見出し{c_idx}]]"] = headers[c_idx - 1] if len(headers) >= c_idx else ""
                    for r_idx in range(1, 4):
                        row = rows[r_idx - 1] if len(rows) >= r_idx else []
                        for c_idx in range(1, 5):
                            ch_map[f"[[行{r_idx}列{c_idx}]]"] = row[c_idx - 1] if (isinstance(row, list) and len(row) >= c_idx) else ""

                    process_slide_shapes(
                        ch_slide, ch_map,
                        chapter_num=ch_num_str, chapter_title=clean_ch_title,
                        is_cover=False
                    )

                    if target_template_idx == 5:
                        build_perfect_steps(ch_slide, step_descs, body_text_val)
                    elif target_template_idx == 9:
                        format_checklist_slide(ch_slide)
                    elif target_template_idx in [6, 7, 8]:
                        format_table_and_note(ch_slide)
                    elif should_draw_chart:
                        format_chart_slide_left_column(ch_slide)
                    elif target_template_idx == 2:
                        format_image_slide_layout(ch_slide, assigned_imgs, captions)
                        if len(assigned_imgs) > 0:
                            for af in assigned_imgs:
                                inserted_image_report.append(f"スライド {ch_page_num}（{clean_ch_title}）: 『{af['filename']}』")

                    if should_draw_chart:
                        c_dict = chart_info if (isinstance(chart_info, dict) and chart_info.get("values")) else extract_chart_data_from_memo(raw_memo)
                        categories = c_dict.get("categories", ["4月", "5月", "6月", "7月", "8月", "9月"])
                        series_name = str(c_dict.get("series_name", "売上高（万円）"))
                        raw_values = c_dict.get("values", [1200.0, 1450.0, 1380.0, 1600.0, 1850.0, 1720.0])

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
                                c_type = XL_CHART_TYPE.COLUMN_CLUSTERED if ("bar" in layout_type or "棒" in clean_ch_title) else XL_CHART_TYPE.LINE_MARKERS
                                x_pos, y_pos, cx_pos, cy_pos = Pt(380), Pt(135), Pt(290), Pt(260)
                                chart_shape = ch_slide.shapes.add_chart(c_type, x_pos, y_pos, cx_pos, cy_pos, chart_data)
                                chart = chart_shape.chart
                                chart.has_title = True
                                chart.chart_title.text_frame.text = str(c_dict.get("title", "月別売上推移（万円）"))

                                for p in chart.chart_title.text_frame.paragraphs:
                                    p.font.name = "メイリオ"
                                    p.font.size = Pt(11)

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

                                force_meiryo_and_style_on_chart(chart, font_size_pt=9)
                                chart_created = True
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

                # ★【最重要】PowerPointの修復警告を完全根絶する整合性修正★
                fix_all_xml_integrity_safe(prs)

                # 6. 保存
                safe_filename = re.sub(r'[/\\:*?"<>|]', '_', doc_title.strip())
                if not safe_filename:
                    safe_filename = "提案書"
                output_path = f"{safe_filename}.pptx"

                prs.save(output_path)

                st.balloons()
                st.success(f"🎉 『{safe_filename}.pptx』の自動構築が完了しました！（全 {len(prs.slides)} ページ）")
                
                if inserted_image_report:
                    st.info("📷 **以下のスライドに画像を配置しました：**\n- " + "\n- ".join(inserted_image_report))

                with open(output_path, "rb") as f:
                    st.download_button(
                        label=f"📥 『{safe_filename}.pptx』をダウンロード",
                        data=f,
                        file_name=output_path,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            except Exception as e:
                st.error(f"エラーが発生しました: {e}")
