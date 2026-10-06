import streamlit as st
from google import genai
from pptx import Presentation
from pptx.util import Pt
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
import unicodedata

st.set_page_config(page_title="W3 AutoSlide", layout="wide")

# カスタムCSS（デザイン調整）
st.markdown("""
    <style>
    img {
        border-radius: 0px !important;
    }
    div.stButton > button {
        font-weight: bold !important;
        border-radius: 8px !important;
        font-size: 16px !important;
    }
    .main-btn > button {
        background-color: #ff4b4b !important;
        color: #ffffff !important;
        font-size: 18px !important;
        padding: 10px 24px !important;
    }
    .main-btn > button:hover {
        background-color: #d93025 !important;
        color: #ffffff !important;
    }
    .edit-card {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #e9ecef;
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

st.image("logo.png", width=350)
st.caption("AI資料 自動生成システム（プレビュー＆自動調整機能付き）")

# ==========================================
# API呼び出し用リトライエンジン（503/429混雑対策）
# ==========================================
def call_gemini_with_retry(client, prompt, max_retries=5):
    """混雑（503/429）発生時に数秒待機して自動リトライする安全ロジック"""
    model_name = 'gemini-3.6-flash'
    last_exception = None

    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt
            )
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

# ==========================================
# クリーニング ＆ サニタイズ
# ==========================================
def clean_notice_text(text):
    if not text:
        return ""
    return re.sub(r'^(?:[※\*\:\s]|注意[：:]|注[：:])+', '', str(text).strip())

def sanitize_main_item(main_item, clean_ch_title):
    """主要項目（赤帯）が15文字を超える長文の場合、自動的に短縮"""
    main_item = str(main_item).strip() if main_item else ""
    if not main_item or len(main_item) > 18 or "。" in main_item or "、" in main_item:
        first_part = re.split(r'[。\n、,]', main_item)[0].strip()
        if 2 <= len(first_part) <= 15:
            main_item = first_part
        else:
            if "推移" in clean_ch_title or "売上" in clean_ch_title:
                main_item = "売上高は堅調に拡大"
            elif "要因" in clean_ch_title or "分析" in clean_ch_title:
                main_item = "増収要因と一時的影響"
            elif "アクション" in clean_ch_title or "手順" in clean_ch_title:
                main_item = "今後の4つの実行計画"
            else:
                main_item = clean_ch_title[:12]
    return main_item

# ==========================================
# 機能③: 【文字溢れ防止】自動フォントサイズ調整エンジン
# ==========================================
def auto_fit_font_size(paragraph, text, base_size_pt=14, min_size_pt=9.5):
    """本文や長文テキストの文字数・改行数に応じてフォントサイズを動的に自動調整"""
    char_count = len(text)
    line_count = text.count('\n') + 1

    target_size = base_size_pt
    if char_count > 150 or line_count >= 7:
        target_size = max(min_size_pt, base_size_pt - 4.5)
    elif char_count > 100 or line_count >= 5:
        target_size = max(min_size_pt, base_size_pt - 3.0)
    elif char_count > 60 or line_count >= 3:
        target_size = max(min_size_pt, base_size_pt - 1.5)

    p_size = Pt(target_size)
    if len(paragraph.runs) > 0:
        for run in paragraph.runs:
            run.font.size = p_size
    else:
        paragraph.font.size = p_size

    return target_size

# ==========================================
# メモテキストからの数値自動抽出
# ==========================================
def extract_chart_data_from_memo(text):
    pattern = r'(\d{1,2}月)[^\d\n]{1,10}?(\d+(?:\.\d+)?)\s*(?:万|万円)?'
    matches = re.findall(pattern, text)
    if matches:
        categories = []
        values = []
        for m in matches:
            cat = m[0]
            val = float(m[1])
            if cat not in categories:
                categories.append(cat)
                values.append(val)
        if len(categories) >= 2:
            return {
                "title": "直近4ヶ月の売上推移",
                "categories": categories,
                "series_name": "売上高（万円）",
                "values": values
            }
    return {
        "title": "直近4ヶ月の売上推移",
        "categories": ["7月", "8月", "9月", "10月"],
        "series_name": "売上高（万円）",
        "values": [100.0, 150.0, 130.0, 200.0]
    }

# ==========================================
# 安全・完全独立なスライド複製クローンエンジン
# ==========================================
def duplicate_slide_safe(prs, src_idx):
    src_slide = prs.slides[src_idx]
    new_slide = prs.slides.add_slide(src_slide.slide_layout)
    
    spTree = new_slide.shapes._spTree
    for child in list(spTree):
        tag = child.tag.split('}')[-1]
        if tag in ['sp', 'graphicFrame', 'grpSp', 'cxnSp', 'pic']:
            spTree.remove(child)
            
    max_id = 2000
    for sld in prs.slides:
        for node in sld.shapes._spTree.iter():
            if 'id' in node.attrib and node.attrib['id'].isdigit():
                val = int(node.attrib['id'])
                if val > max_id:
                    max_id = val

    rel_map = {}
    for shape in src_slide.shapes:
        new_el = copy.deepcopy(shape._element)
        for node in new_el.iter():
            if 'id' in node.attrib and node.attrib['id'].isdigit():
                max_id += 1
                node.attrib['id'] = str(max_id)
                
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

# ==========================================
# ページ番号確実配置エンジン
# ==========================================
def ensure_page_number(slide, page_num):
    found_page_shape = None
    for shape in slide.shapes:
        is_placeholder = (shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER)
        is_name_match = any(kw in shape.name.lower() for kw in ["スライド番号", "page_number", "pagenumber", "slide_number"])
        is_bottom_right_digit = False
        if shape.has_text_frame:
            t = shape.text_frame.text.strip()
            if t.isdigit() and shape.top > 6000000 and shape.left > 6500000:
                is_bottom_right_digit = True

        if is_placeholder or is_name_match or is_bottom_right_digit:
            found_page_shape = shape
            break

    if found_page_shape:
        tf = found_page_shape.text_frame
        if len(tf.paragraphs) > 0 and len(tf.paragraphs[0].runs) > 0:
            tf.paragraphs[0].runs[0].text = str(page_num)
            for r in tf.paragraphs[0].runs[1:]:
                r.text = ""
        else:
            tf.text = str(page_num)
    else:
        txBox = slide.shapes.add_textbox(7922128, 6960925, 2405062, 401638)
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = str(page_num)
        p.font.size = Pt(12)

# ==========================================
# グラフフォント（メイリオ ＋ 9pt小ぶり化）XMLレベル強制適用
# ==========================================
def force_meiryo_and_size_on_chart(chart, font_size_pt=9):
    for axis in [getattr(chart, 'category_axis', None), getattr(chart, 'value_axis', None)]:
        if axis and hasattr(axis, 'tick_labels'):
            try:
                axis.tick_labels.font.name = "メイリオ"
                axis.tick_labels.font.size = Pt(font_size_pt)
                txPr = axis.tick_labels._element.get_or_add_txPr()
                for p in txPr.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p'):
                    pPr = p.get_or_add_pPr()
                    defRPr = pPr.get_or_add_defRPr()
                    defRPr.attrib['sz'] = str(font_size_pt * 100)
                    defRPr.attrib['typeface'] = 'メイリオ'
                    for child in list(defRPr):
                        if child.tag.endswith('ea') or child.tag.endswith('latin'):
                            defRPr.remove(child)
                    defRPr.append(parse_xml(f'<a:ea {nsdecls("a")} typeface="メイリオ"/>'))
                    defRPr.append(parse_xml(f'<a:latin {nsdecls("a")} typeface="Meiryo"/>'))
            except Exception:
                pass

    if chart.has_legend and chart.legend:
        try:
            chart.legend.font.name = "メイリオ"
            chart.legend.font.size = Pt(font_size_pt)
            txPr = chart.legend._element.get_or_add_txPr()
            for p in txPr.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p'):
                pPr = p.get_or_add_pPr()
                defRPr = pPr.get_or_add_defRPr()
                defRPr.attrib['sz'] = str(font_size_pt * 100)
                defRPr.append(parse_xml(f'<a:ea {nsdecls("a")} typeface="メイリオ"/>'))
                defRPr.append(parse_xml(f'<a:latin {nsdecls("a")} typeface="Meiryo"/>'))
        except Exception:
            pass

# ==========================================
# グラフ追加スライドのレイアウト自動調整
# ==========================================
def adjust_shapes_for_chart(slide):
    for shape in slide.shapes:
        if "SECTION_LINE" in shape.name or "LINE" in shape.name.upper():
            shape.width = Pt(4.0)
            shape.height = Pt(19.3)
        elif shape.has_text_frame and shape.top > Pt(80) and Pt(100) <= shape.left < Pt(350):
            shape.width = Pt(220)

# ==========================================
# タイトル段落の番号保持・確実置換
# ==========================================
def format_chapter_title_paragraph(p, ch_num, ch_title):
    clean_title = re.sub(r'^\d+[\.\s_]*', '', str(ch_title)).strip() if ch_title else ""
    full_text = p.text
    if "00" in full_text or "[[章タイトル]]" in full_text:
        new_text = full_text
        new_text = re.sub(r'00\s*\[\[章タイトル\]\]', f"{ch_num} {clean_title}", new_text)
        new_text = re.sub(r'\[\[章タイトル\]\]', clean_title, new_text)
        new_text = re.sub(r'\b00\b', ch_num, new_text)
        if not new_text.startswith(ch_num):
            new_text = f"{ch_num} {new_text}"

        if len(p.runs) > 0:
            p.runs[0].text = new_text
            for r in p.runs[1:]:
                r.text = ""
        else:
            p.text = new_text

# ==========================================
# 段落置換処理（自動フォントアジャスト付き）
# ==========================================
def process_paragraph_runs(p, replace_map, chapter_num=None, chapter_title=None):
    if "00" in p.text or "[[章タイトル]]" in p.text:
        format_chapter_title_paragraph(p, chapter_num if chapter_num else "01", chapter_title)

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

    # 本文テキストの文字数に応じた自動フォントサイズ調整
    if p_text and len(p_text.strip()) > 30 and not any(tag in p_text for tag in ["[[章タイトル]]", "STEP", "P1", "P2"]):
        auto_fit_font_size(p, p_text, base_size_pt=14, min_size_pt=9.5)

# ==========================================
# スライド要素置換処理
# ==========================================
def process_slide_shapes(slide, replace_map, chapter_num=None, chapter_title=None, assigned_files=None, captions=None):
    shapes_to_remove = []
    has_assigned = assigned_files is not None and len(assigned_files) > 0

    notice_val = replace_map.get("[[注意文]]", "") or replace_map.get("[[補足文]]", "")
    has_notice = bool(notice_val and str(notice_val).strip())

    sub_item_val = replace_map.get("[[補助項目]]", "") or replace_map.get("[[補助項目の本文]]", "")
    has_sub_item = bool(sub_item_val and str(sub_item_val).strip())

    body_val = replace_map.get("[[本文]]", "")
    has_body = bool(body_val and str(body_val).strip())

    for shape in slide.shapes:
        sname_upper = shape.name.upper()
        raw_text = shape.text_frame.text if shape.has_text_frame else ""

        if shape.has_text_frame:
            txt = raw_text.strip()
            if txt.isdigit() and len(txt) <= 2:
                is_bottom_right = (shape.top > Pt(400) and shape.left > Pt(500))
                is_placeholder = (shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER)
                if not is_bottom_right and not is_placeholder:
                    shapes_to_remove.append(shape)
                    continue

        if not has_notice:
            if "CAUTION" in sname_upper or "SUPPLEMENT" in sname_upper or "補足" in sname_upper or "注意" in sname_upper:
                shapes_to_remove.append(shape)
                continue
            if shape.has_text_frame:
                if any(tag in raw_text for tag in ["[[注意文]]", "[[補足文]]", "[[表の注記]]"]):
                    shapes_to_remove.append(shape)
                    continue
                if raw_text.strip() in ["注意", "補足", "※", "注"]:
                    shapes_to_remove.append(shape)
                    continue

        if not has_sub_item and shape.has_text_frame:
            if "[[補助項目]]" in raw_text or "[[補助項目の本文]]" in raw_text:
                shapes_to_remove.append(shape)
                continue

        if not has_body and shape.has_text_frame:
            if "[[本文]]" in raw_text:
                shapes_to_remove.append(shape)
                continue

        if shape.has_text_frame and ("[[画像1]]" in shape.text_frame.text or "[[画像2]]" in shape.text_frame.text):
            tf_text = shape.text_frame.text
            target_idx = 0 if "[[画像1]]" in tf_text else 1

            if has_assigned and len(assigned_files) > target_idx:
                img_file = assigned_files[target_idx]
                img_bytes_data = img_file.getvalue()

                pil_img = Image.open(io.BytesIO(img_bytes_data))
                orig_w, orig_h = pil_img.size

                box_w, box_h = shape.width, shape.height
                box_left, box_top = shape.left, shape.top

                scale = min(box_w / orig_w, box_h / orig_h)
                new_w = int(orig_w * scale)
                new_h = int(orig_h * scale)
                new_left = int(box_left + (box_w - new_w) / 2)
                new_top = int(box_top + (box_h - new_h) / 2)

                slide.shapes.add_picture(io.BytesIO(img_bytes_data), new_left, new_top, new_w, new_h)

            shapes_to_remove.append(shape)
            continue

        if shape.has_text_frame:
            tf = shape.text_frame

            if "[[キャプション1]]" in raw_text:
                if not has_assigned or len(assigned_files) < 1:
                    shapes_to_remove.append(shape)
                    continue
                elif captions and len(captions) > 0:
                    replace_map["[[キャプション1]]"] = captions[0]

            if "[[キャプション2]]" in raw_text:
                if not has_assigned or len(assigned_files) < 2:
                    shapes_to_remove.append(shape)
                    continue
                elif captions and len(captions) > 1:
                    replace_map["[[キャプション2]]"] = captions[1]

            tf.word_wrap = True
            for p in tf.paragraphs:
                if "[[チェック項目" in p.text:
                    for k in range(1, 7):
                        tag_k = f"[[チェック項目{k}]]"
                        if tag_k in p.text and replace_map.get(tag_k, "").strip() == "":
                            p.text = ""

                process_paragraph_runs(p, replace_map, chapter_num, chapter_title)

        if shape.has_table:
            for cell in shape.table.iter_cells():
                for p in cell.text_frame.paragraphs:
                    process_paragraph_runs(p, replace_map)

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
        st.rerun()

col1, col2 = st.columns(2)
with col1:
    client_name = st.text_input("提案先（顧客名）", "例）役員会 各位")
    doc_title = st.text_input("資料タイトル", "例）売上推移と今後の成長戦略")
with col2:
    sender_name = st.text_input("作成者・部署など", "例）営業部")
    uploaded_files = st.file_uploader("挿入画像（任意・複数可）", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

image_contexts = []
if uploaded_files:
    st.write("📷 画像の用途・説明")
    c_cols = st.columns(min(len(uploaded_files), 3))
    for i, f in enumerate(uploaded_files):
        with c_cols[i % 3]:
            desc = st.text_input(f"画像{i+1}（{f.name}）の説明", placeholder="例：サービス活用イメージ図")
            image_contexts.append({"index": i, "filename": f.name, "description": desc})

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
4. 月商200万円の継続的維持""", height=220)


# ==========================================
# ステップ1: AIで構造解析（プレビュー生成）
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

                    【レイアウト＆情報抽出ルール】
                    1. メモの主要項目を整理し、プレゼンに最適な3〜7個の章（chapters）にまとめてください。
                    2. 各章のコンテンツに合わせて、最適な "layout_type" を以下から選定してください:
                       - "line_chart": 月別推移、売上推移、時系列トレンドなどの数値変化がある場合
                       - "bar_chart": 項目別比較、実績数値の比較、カテゴリ別実績がある場合
                       - "text": 概要・課題・特徴・メリット・期待効果・要因分析など
                       - "step": 導入手順・運用フロー・タイムライン・実行ステップ（最大4ステップ）
                       - "checklist": 確認事項・必要書類・チェックリストなど
                       - "table": 料金プランや機能比較などのテキスト表
                    3. "main_item" について:
                       - 赤帯に配置する15文字以内の短いキャッチコピー（見出し）を指定してください。
                    4. "step_descs" について:
                       - layout_type が "step" の場合、メモから具体的なアクションやステップを4つ抽出し、"step_descs" 配列に格納してください。
                    5. "chart_info" について:
                       - メモ内に売上・件数・数値・月別などのデータがある場合は、必ず "chart_info" に "title", "categories", "series_name", "values" を格納してください。

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
                          "step_descs": [
                            "顧客購買データの詳細分析",
                            "高単価ターゲット層へのアプローチ強化",
                            "新プロモーションの実施",
                            "月商200万円の継続的維持"
                          ],
                          "table": {{ "headers": [], "rows": [] }},
                          "assigned_image_indices": [],
                          "image_captions": [],
                          "chart_info": {{
                            "title": "直近4ヶ月の売上推移",
                            "categories": ["7月", "8月", "9月", "10月"],
                            "series_name": "売上高（万円）",
                            "values": [100, 150, 130, 200]
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
                    st.success("✅ AIの構成解析が完了しました！下部の編集フォームでテキストや数値を微調整できます。")
                except Exception as e:
                    err_msg = str(e)
                    if "503" in err_msg or "UNAVAILABLE" in err_msg:
                        st.error("【サーバー混雑中】現在AIサーバーへのアクセスが集中しています。恐れ入りますが1〜2分待ってから再度お試しください。")
                    elif "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                        st.error("【利用制限】1分あたりの利用上限に達しました。少し待ってから再度お試しください。")
                    else:
                        st.error(f"解析中にエラーが発生しました: {e}")

# ==========================================
# 機能①: WEB編集フォーム ＆ 最終生成
# ==========================================
if st.session_state.get("analyzed", False) and "parsed_data" in st.session_state:
    st.subheader("✏️ AI解析結果のプレビュー ＆ WEB編集フォーム")
    st.caption("各スライドの内容を自由に変更できます。調整が終わったら最下部の生成ボタンを押してください。")

    parsed_chapters = st.session_state["parsed_data"].get("chapters", [])
    edited_chapters = []

    for idx, ch in enumerate(parsed_chapters):
        with st.expander(f"📄 スライド {idx + 1}: {ch.get('chapter_title', '')}", expanded=(idx == 0)):
            ec_cols1, ec_cols2 = st.columns([2, 1])
            
            with ec_cols1:
                e_title = st.text_input(f"スライド{idx+1} タイトル", value=ch.get("chapter_title", ""), key=f"title_{idx}")
                e_main = st.text_input(f"スライド{idx+1} 主要見出し（赤帯・15文字以内）", value=ch.get("main_item", ""), key=f"main_{idx}")
                e_body = st.text_area(f"スライド{idx+1} 本文", value=ch.get("body", ""), height=100, key=f"body_{idx}")

            with ec_cols2:
                layout_options = ["line_chart", "bar_chart", "text", "step", "checklist", "table"]
                current_layout = str(ch.get("layout_type", "text")).lower()
                l_idx = layout_options.index(current_layout) if current_layout in layout_options else 2
                e_layout = st.selectbox(f"レイアウトタイプ", layout_options, index=l_idx, key=f"layout_{idx}")
                
                e_sub_title = st.text_input(f"補助項目 タイトル", value=ch.get("sub_title", ""), key=f"sub_t_{idx}")
                e_sub_body = st.text_area(f"補助項目 本文", value=ch.get("sub_body", ""), height=68, key=f"sub_b_{idx}")

            if e_layout == "step":
                st.markdown("**ステップ詳細（最大4ステップ）**")
                s_descs = ch.get("step_descs", [])
                e_s_descs = []
                s_cols = st.columns(2)
                for s_i in range(4):
                    default_s = s_descs[s_i] if len(s_descs) > s_i else ""
                    with s_cols[s_i % 2]:
                        val_s = st.text_input(f"STEP {s_i+1}", value=default_s, key=f"step_{idx}_{s_i}")
                        if val_s: e_s_descs.append(val_s)
            else:
                e_s_descs = ch.get("step_descs", [])

            c_info = ch.get("chart_info") or extract_chart_data_from_memo(raw_memo)
            if e_layout in ["line_chart", "bar_chart"] or idx == 0:
                st.markdown("**📊 グラフデータ設定**")
                g_cols = st.columns(3)
                with g_cols[0]:
                    g_title = st.text_input("グラフタイトル", value=c_info.get("title", "売上推移"), key=f"gt_{idx}")
                with g_cols[1]:
                    g_cats_str = st.text_input("カテゴリ（カンマ区切り）", value=",".join(map(str, c_info.get("categories", ["7月","8月","9月","10月"]))), key=f"gc_{idx}")
                with g_cols[2]:
                    g_vals_str = st.text_input("数値データ（カンマ区切り）", value=",".join(map(str, c_info.get("values", [100, 150, 130, 200]))), key=f"gv_{idx}")

                cats_list = [c.strip() for c in g_cats_str.split(",") if c.strip()]
                vals_list = []
                for v in g_vals_str.split(","):
                    try:
                        vals_list.append(float(re.sub(r'[^\d\.]', '', v)))
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
                "body": e_body,
                "sub_title": e_sub_title,
                "sub_body": e_sub_body,
                "notice": ch.get("notice", ""),
                "step_descs": e_s_descs,
                "table": ch.get("table", {}),
                "assigned_image_indices": ch.get("assigned_image_indices", []),
                "image_captions": ch.get("image_captions", []),
                "chart_info": edited_c_info
            })

    st.markdown("---")
    
    st.markdown('<div class="main-btn">', unsafe_allow_html=True)
    if st.button("✨ この内容で PowerPoint 資料を自動生成する ✨", use_container_width=True):
        status_box = st.empty()
        with st.spinner("WEB編集結果を元にPowerPoint資料を構築中..."):
            try:
                current_dir = os.path.dirname(os.path.abspath(__file__))
                template_path = os.path.join(current_dir, '資料作成テンプレート_A4横.pptx')

                if not os.path.exists(template_path):
                    all_files = os.listdir(current_dir)
                    target_norm = unicodedata.normalize('NFC', '資料作成テンプレート_A4横.pptx')
                    for f in all_files:
                        if f.endswith('.pptx'):
                            if unicodedata.normalize('NFC', f) == target_norm or "テンプレート" in f:
                                template_path = os.path.join(current_dir, f)
                                break
                    else:
                        pptx_list = [f for f in all_files if f.endswith('.pptx')]
                        if pptx_list:
                            template_path = os.path.join(current_dir, pptx_list[0])
                        else:
                            raise FileNotFoundError(f"テンプレートファイルが見つかりません。")

                prs = Presentation(template_path)
                tpl_slide_count = len(prs.slides)
                file_name_only = os.path.basename(template_path)

                today_str = datetime.date.today().strftime("%Y/%m/%d")
                new_created_slides = []

                # 1. 表紙スライド
                cover_slide = duplicate_slide_safe(prs, 0)
                cover_map = {
                    "[[資料タイトル]]": doc_title,
                    "[[顧客名]]": client_name,
                    "[[バージョン]]": "1.0",
                    "[[更新日]]": today_str
                }
                process_slide_shapes(cover_slide, cover_map)
                new_created_slides.append(cover_slide)

                # 2. INDEX（目次）スライド
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

                for k in range(len(edited_chapters) + 1, 7):
                    index_map[f"[[章{k}タイトル]]"] = ""
                    index_map[f"P[[章{k}ページ]]"] = ""
                    index_map[f"[[章{k}ページ]]"] = ""

                process_slide_shapes(index_slide, index_map)
                new_created_slides.append(index_slide)

                chart_created = False

                # 3. 各章スライド構築
                for ch_idx, ch in enumerate(edited_chapters):
                    ch_num_str = f"{ch_idx + 1:02d}"
                    ch_page_num = ch_idx + 2
                    ch_title = ch.get("chapter_title", f"章{ch_idx+1}")
                    clean_ch_title = re.sub(r'^\d+[\.\s_]*', '', str(ch_title)).strip()

                    layout_type = str(ch.get("layout_type", "text")).lower()
                    img_indices = ch.get("assigned_image_indices", [])
                    img_count = len(img_indices)

                    assigned_files = [uploaded_files[i] for i in img_indices if uploaded_files and i < len(uploaded_files)]
                    captions = ch.get("image_captions", [])

                    chart_info = ch.get("chart_info")
                    has_values = bool(chart_info and isinstance(chart_info.get("values"), list) and len(chart_info.get("values")) > 0)
                    is_chart_keyword = any(kw in clean_ch_title for kw in ["推移", "売上", "実績", "比較", "グラフ", "データ"])

                    should_draw_chart = (not chart_created) and (
                        "line" in layout_type or "bar" in layout_type or "chart" in layout_type or
                        is_chart_keyword or ch_idx == 0
                    )

                    target_template_idx = 2
                    if layout_type == "checklist" or "確認" in clean_ch_title or "チェック" in clean_ch_title:
                        target_template_idx = 9 if tpl_slide_count > 9 else 2
                    elif layout_type == "step" or "フロー" in clean_ch_title or "流れ" in clean_ch_title or "手順" in clean_ch_title or "アクション" in clean_ch_title:
                        target_template_idx = 5 if tpl_slide_count > 5 else 2
                    elif layout_type == "table" or "プラン" in clean_ch_title or "料金" in clean_ch_title:
                        tbl_data = ch.get("table", {})
                        cols = ch.get("column_count", len(tbl_data.get("headers", [])))
                        if cols == 4 and tpl_slide_count > 8: target_template_idx = 8
                        elif cols == 3 and tpl_slide_count > 7: target_template_idx = 7
                        elif tpl_slide_count > 6: target_template_idx = 6
                        else: target_template_idx = 2
                    else:
                        if img_count >= 2 and tpl_slide_count > 3: target_template_idx = 3
                        elif img_count == 1 and tpl_slide_count > 4: target_template_idx = 4
                        else: target_template_idx = 2

                    if target_template_idx >= tpl_slide_count:
                        target_template_idx = 2

                    ch_slide = duplicate_slide_safe(prs, target_template_idx)
                    new_created_slides.append(ch_slide)

                    main_item_text = sanitize_main_item(ch.get("main_item", ""), clean_ch_title)

                    ch_map = {}
                    ch_map["[[主要項目]]"] = main_item_text
                    ch_map["[[本文]]"] = ch.get("body", "")

                    if should_draw_chart:
                        ch_map["[[補助項目]]"] = ""
                        ch_map["[[補助項目の本文]]"] = ""
                    else:
                        ch_map["[[補助項目]]"] = ch.get("sub_title", "実行ステップ")
                        ch_map["[[補助項目の本文]]"] = ch.get("sub_body", ch.get("body", ""))

                    notice_val = clean_notice_text(ch.get("notice", ""))
                    ch_map["[[補足文]]"] = notice_val
                    ch_map["[[注意文]]"] = notice_val

                    step_descs = ch.get("step_descs", [])
                    for idx_s in range(1, 5):
                        desc_s = step_descs[idx_s - 1] if len(step_descs) >= idx_s else ""
                        ch_map[f"[[STEP{idx_s}の説明]]"] = desc_s

                    process_slide_shapes(
                        ch_slide, ch_map,
                        chapter_num=ch_num_str, chapter_title=clean_ch_title,
                        assigned_files=assigned_files, captions=captions
                    )

                    if should_draw_chart and chart_info:
                        categories = chart_info.get("categories", ["7月", "8月", "9月", "10月"])
                        series_name = str(chart_info.get("series_name", "売上高（万円）"))
                        raw_values = chart_info.get("values", [100.0, 150.0, 130.0, 200.0])

                        values = []
                        for v in raw_values:
                            try:
                                num_str = re.sub(r'[^\d\.]', '', str(v))
                                values.append(float(num_str) if num_str else 0.0)
                            except Exception:
                                values.append(0.0)

                        if categories and values:
                            min_len = min(len(categories), len(values))
                            categories = categories[:min_len]
                            values = values[:min_len]

                            chart_data = CategoryChartData()
                            chart_data.categories = categories
                            chart_data.add_series(series_name, values)

                            adjust_shapes_for_chart(ch_slide)

                            try:
                                c_type = XL_CHART_TYPE.COLUMN_CLUSTERED if ("bar" in layout_type or "棒" in clean_ch_title) else XL_CHART_TYPE.LINE_MARKERS
                                x_pos, y_pos, cx_pos, cy_pos = Pt(380), Pt(135), Pt(290), Pt(260)
                                chart_shape = ch_slide.shapes.add_chart(c_type, x_pos, y_pos, cx_pos, cy_pos, chart_data)
                                chart = chart_shape.chart

                                chart.has_title = False
                                force_meiryo_and_size_on_chart(chart, font_size_pt=9)
                                chart_created = True
                            except Exception:
                                pass

                # 4. 元のテンプレートスライド消去
                for i in range(tpl_slide_count - 1, -1, -1):
                    del prs.slides._sldIdLst[i]

                # 5. 残存未置換タグ消去 ＆ ページ番号付与
                for s_i, sld in enumerate(prs.slides):
                    for shp in sld.shapes:
                        if shp.has_text_frame:
                            for p in shp.text_frame.paragraphs:
                                if "[[" in p.text and "]]" in p.text:
                                    p.text = re.sub(r'\[\[.*?\]\]', '', p.text)
                    if s_i > 0:
                        ensure_page_number(sld, s_i)

                # 6. 保存
                safe_filename = re.sub(r'[\\/:*?"<>|]', '_', doc_title.strip())
                if not safe_filename:
                    safe_filename = "提案書"
                output_path = f"{safe_filename}.pptx"

                prs.save(output_path)

                st.balloons()
                st.success(f"🎉 『{safe_filename}.pptx』の自動構築が完了しました！")
                with open(output_path, "rb") as f:
                    st.download_button(
                        label=f"📥 『{safe_filename}.pptx』をダウンロード",
                        data=f,
                        file_name=output_path,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            except Exception as e:
                st.error(f"エラーが発生しました: {e}")
    st.markdown('</div>', unsafe_allow_html=True)
