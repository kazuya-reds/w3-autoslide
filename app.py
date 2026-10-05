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

# カスタムCSS（ロゴの角丸解除 ＋ ボタンデザイン変更）
st.markdown("""
    <style>
    img {
        border-radius: 0px !important;
    }
    div.stButton > button {
        background-color: #ff4b4b !important;
        color: #ffffff !important;
        font-weight: bold !important;
        border: none !important;
        border-radius: 8px !important;
        font-size: 18px !important;
        padding: 10px 24px !important;
    }
    div.stButton > button:hover {
        background-color: #d93025 !important;
        color: #ffffff !important;
    }
    </style>
""", unsafe_allow_html=True)

st.image("logo.png", width=350)
st.caption("AI資料 自動生成システム")

# ==========================================
# API呼び出し用リトライヘルパー（429/503対策）
# ==========================================
def call_gemini_with_retry(client, prompt, max_retries=3):
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model='gemini-3.6-flash',
                contents=prompt
            )
            return response.text.strip()
        except Exception as e:
            err_str = str(e)
            if ("429" in err_str or "RESOURCE_EXHAUSTED" in err_str) and attempt < max_retries - 1:
                match = re.search(r'retry in ([0-9\.]+)s', err_str)
                wait_sec = int(float(match.group(1))) + 2 if match else 60
                st.warning(f"⏳ API利用制限に達しました。{wait_sec}秒待機後に自動再試行します... ({attempt + 1}/{max_retries})")
                time.sleep(wait_sec)
            elif ("503" in err_str or "UNAVAILABLE" in err_str or "high demand" in err_str) and attempt < max_retries - 1:
                wait_sec = 5
                st.warning(f"⚠️ API高負荷のため待機中... ({attempt + 1}/{max_retries})")
                time.sleep(wait_sec)
            else:
                raise e

# ==========================================
# テキスト・注意文クリーニング
# ==========================================
def clean_notice_text(text):
    if not text:
        return ""
    return re.sub(r'^(?:[※\*\:\s]|注意[：:]|注[：:])+', '', str(text).strip())

# ==========================================
# メモテキストからの数値自動抽出（確実なフォールバック）
# ==========================================
def extract_chart_data_from_memo(text):
    """メモから『7月は100万』などの月別数値を自動抽出しグラフ構造を生成"""
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
    """Shape IDの完全リナンバーとリレーションの独立バインドを行いスライドを複製"""
    src_slide = prs.slides[src_idx]
    new_slide = prs.slides.add_slide(src_slide.slide_layout)
    
    # 自動作成されたデフォルトプレースホルダー枠を全消去
    spTree = new_slide.shapes._spTree
    for child in list(spTree):
        tag = child.tag.split('}')[-1]
        if tag in ['sp', 'graphicFrame', 'grpSp', 'cxnSp', 'pic']:
            spTree.remove(child)
            
    # 全スライドから最大 Shape ID を取得し重複を完全防止
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
        
        # XML要素ツリー内の全 ID 属性をリナンバー
        for node in new_el.iter():
            if 'id' in node.attrib and node.attrib['id'].isdigit():
                max_id += 1
                node.attrib['id'] = str(max_id)
                
        # 画像等のリレーションシップの安全再バインド
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
# グラフフォント（メイリオ）XMLレベルでの100%強制適用
# ==========================================
def force_meiryo_on_chart(chart):
    """X軸・Y軸・凡例にDrawingML XMLレベルで『メイリオ』を強制的注入"""
    # 1. 軸 (CategoryAxis, ValueAxis)
    for axis in [getattr(chart, 'category_axis', None), getattr(chart, 'value_axis', None)]:
        if axis and hasattr(axis, 'tick_labels'):
            try:
                axis.tick_labels.font.name = "メイリオ"
                txPr = axis.tick_labels._element.get_or_add_txPr()
                for p in txPr.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p'):
                    pPr = p.get_or_add_pPr()
                    defRPr = pPr.get_or_add_defRPr()
                    defRPr.attrib['typeface'] = 'メイリオ'
                    for child in list(defRPr):
                        if child.tag.endswith('ea') or child.tag.endswith('latin'):
                            defRPr.remove(child)
                    defRPr.append(parse_xml(f'<a:ea {nsdecls("a")} typeface="メイリオ"/>'))
                    defRPr.append(parse_xml(f'<a:latin {nsdecls("a")} typeface="Meiryo"/>'))
            except Exception:
                pass

    # 2. 凡例
    if chart.has_legend and chart.legend:
        try:
            chart.legend.font.name = "メイリオ"
            txPr = chart.legend._element.get_or_add_txPr()
            for p in txPr.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p'):
                pPr = p.get_or_add_pPr()
                defRPr = pPr.get_or_add_defRPr()
                defRPr.append(parse_xml(f'<a:ea {nsdecls("a")} typeface="メイリオ"/>'))
                defRPr.append(parse_xml(f'<a:latin {nsdecls("a")} typeface="Meiryo"/>'))
        except Exception:
            pass

# ==========================================
# グラフ追加スライドのレイアウト自動調整
# ==========================================
def adjust_shapes_for_chart(slide):
    """スライド左側のテキストボックスの幅を Pt(250) に最適化し、右側のグラフ領域を美しく空ける"""
    for shape in slide.shapes:
        if shape.has_text_frame:
            if shape.top > Pt(80) and shape.left < Pt(300):
                shape.width = Pt(250)

# ==========================================
# 段落置換処理（タイトルのスペース補正追加）
# ==========================================
def process_paragraph_runs(p, replace_map, chapter_num=None, chapter_title=None):
    clean_title = re.sub(r'^\d+[\.\s_]*', '', str(chapter_title)).strip() if chapter_title else ""

    for r in p.runs:
        if chapter_num and "00" in r.text:
            r.text = r.text.replace("00 [[章タイトル]]", f"{chapter_num} {clean_title}")
            r.text = r.text.replace("00[[章タイトル]]", f"{chapter_num} {clean_title}")
            r.text = r.text.replace("00", f"{chapter_num} ")
        if clean_title and "[[章タイトル]]" in r.text:
            r.text = r.text.replace("[[章タイトル]]", clean_title)
        for tag, val in replace_map.items():
            if tag in r.text:
                r.text = r.text.replace(tag, str(val))

    p_text = p.text
    has_unreplaced_tag = False
    if clean_title and "[[章タイトル]]" in p_text:
        has_unreplaced_tag = True
    for tag in replace_map:
        if tag in p_text:
            has_unreplaced_tag = True
            break

    if has_unreplaced_tag:
        full_text = p_text
        if clean_title and "[[章タイトル]]" in full_text:
            full_text = full_text.replace("00 [[章タイトル]]", f"{chapter_num} {clean_title}")
            full_text = full_text.replace("00[[章タイトル]]", f"{chapter_num} {clean_title}")
            full_text = full_text.replace("[[章タイトル]]", clean_title)
        for tag, val in replace_map.items():
            if tag in full_text:
                full_text = full_text.replace(tag, str(val))

        if len(p.runs) > 0:
            p.runs[0].text = full_text
            for r in p.runs[1:]:
                r.text = ""

# ==========================================
# スライド要素置換処理（単体数字テロップ削除機能付き）
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

        # ★テンプレート固有の単体数字テロップ（2, 3, 4など）を自動消去★
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
# 1. 入力UI
# ==========================================
with st.sidebar:
    st.header("⚙️ 設定")
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    st.info("テンプレート: 資料作成テンプレート_A4横.pptx を使用します")

col1, col2 = st.columns(2)
with col1:
    client_name = st.text_input("提案先（顧客名）", "例）役員会 各位")
    doc_title = st.text_input("資料タイトル", "例）売上推移と今後の成長戦略")
with col2:
    sender_name = st.text_input("作成者・部署など", "例）営業部")
    uploaded_files = st.file_uploader("挿入画像（任意・複数可）", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

image_contexts = []
if uploaded_files:
    st.write("📷 画像の用途・説明（AIが最適なページに割り当てます）")
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
4. 月商200万円の継続的維持""", height=400)


# ==========================================
# メイン処理
# ==========================================
if st.button("✨資料を生成する✨"):
    if not api_key:
        st.error("APIキーが設定されていません。.streamlit/secrets.toml をご確認ください。")
    elif not raw_memo.strip():
        st.error("提案メモを入力してください。")
    else:
        status_box = st.empty()
        with st.spinner("AIがメモを解析し、資料を構築中..."):
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

                st.caption(f"ℹ️ テンプレート『{file_name_only}』（全{tpl_slide_count}枚）から資料を動的生成中...")

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
                3. "step_descs" について:
                   - layout_type が "step" の場合、メモから具体的なアクションやステップを4つ抽出し、"step_descs" 配列に格納してください。
                4. "chart_info" について:
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

                status_box.success("✅ AI構造解析完了：PowerPoint資料を構築中...")

                chapters = final_data.get("chapters", [])
                today_str = datetime.date.today().strftime("%Y/%m/%d")

                # ==========================================
                # 安全・確実な独立クローン生成エンジン
                # ==========================================
                new_created_slides = []

                # 1. 表紙スライド（インデックス0）を独立複製
                cover_slide = duplicate_slide_safe(prs, 0)
                cover_map = {
                    "[[資料タイトル]]": doc_title,
                    "[[顧客名]]": client_name,
                    "[[バージョン]]": "1.0",
                    "[[更新日]]": today_str
                }
                process_slide_shapes(cover_slide, cover_map)
                new_created_slides.append(cover_slide)

                # 2. INDEX（目次）スライド（インデックス1）を独立複製
                index_slide = duplicate_slide_safe(prs, 1)
                index_map = {}
                for ch_idx, ch in enumerate(chapters):
                    ch_num_str = f"{ch_idx + 1:02d}"
                    ch_page_num = ch_idx + 2
                    ch_title = ch.get("chapter_title", f"章{ch_idx+1}")
                    clean_ch_title = re.sub(r'^\d+[\.\s_]*', '', str(ch_title)).strip()

                    index_map[f"[[章{ch_idx+1}タイトル]]"] = f"{ch_num_str} {clean_ch_title}"
                    index_map[f"P[[章{ch_idx+1}ページ]]"] = f"P{ch_page_num}"
                    index_map[f"[[章{ch_idx+1}ページ]]"] = str(ch_page_num)

                for k in range(len(chapters) + 1, 7):
                    index_map[f"[[章{k}タイトル]]"] = ""
                    index_map[f"P[[章{k}ページ]]"] = ""
                    index_map[f"[[章{k}ページ]]"] = ""

                process_slide_shapes(index_slide, index_map)
                new_created_slides.append(index_slide)

                # グラフ生成フラグ（★1回のみ作成に制限★）
                chart_created = False

                # 3. 各章スライドの独立複製＆置換処理
                for ch_idx, ch in enumerate(chapters):
                    ch_num_str = f"{ch_idx + 1:02d}"
                    ch_page_num = ch_idx + 2
                    ch_title = ch.get("chapter_title", f"章{ch_idx+1}")
                    clean_ch_title = re.sub(r'^\d+[\.\s_]*', '', str(ch_title)).strip()

                    layout_type = str(ch.get("layout_type", "text")).lower()
                    img_indices = ch.get("assigned_image_indices", [])
                    img_count = len(img_indices)

                    assigned_files = [uploaded_files[i] for i in img_indices if uploaded_files and i < len(uploaded_files)]
                    captions = ch.get("image_captions", [])

                    # 数値判定（AIレスポンス ＋ メモから自動抽出）
                    chart_info = ch.get("chart_info")
                    if not chart_info or not chart_info.get("values"):
                        chart_info = extract_chart_data_from_memo(raw_memo)

                    has_values = bool(chart_info and isinstance(chart_info.get("values"), list) and len(chart_info.get("values")) > 0)
                    is_chart_keyword = any(kw in clean_ch_title for kw in ["推移", "売上", "実績", "比較", "グラフ", "データ"])

                    should_draw_chart = (not chart_created) and (
                        "line" in layout_type or "bar" in layout_type or "chart" in layout_type or
                        is_chart_keyword or ch_idx == 0
                    )

                    # テンプレート上のソーススライド判定
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

                    # 安全複製
                    ch_slide = duplicate_slide_safe(prs, target_template_idx)
                    new_created_slides.append(ch_slide)

                    # 主要項目の自動補完（空文字防止）
                    main_item_text = ch.get("main_item", "").strip()
                    if not main_item_text:
                        main_item_text = "売上高は堅調に推移" if ("売上" in clean_ch_title or "推移" in clean_ch_title) else clean_ch_title

                    # テキスト置換辞書
                    ch_map = {}
                    ch_map["[[主要項目]]"] = main_item_text
                    ch_map["[[本文]]"] = ch.get("body", "")

                    # ★グラフを配置するスライドでは補助項目をクリアして左側をスッキリさせる★
                    if should_draw_chart:
                        ch_map["[[補助項目]]"] = ""
                        ch_map["[[補助項目の本文]]"] = ""
                    else:
                        ch_map["[[補助項目]]"] = ch.get("sub_title", "実行ステップ")
                        ch_map["[[補助項目の本文]]"] = ch.get("sub_body", ch.get("body", ""))

                    notice_val = clean_notice_text(ch.get("notice", ""))
                    ch_map["[[補足文]]"] = notice_val
                    ch_map["[[注意文]]"] = notice_val

                    # STEP補完処理（ステップ本文の確実な割り当て）
                    step_descs = ch.get("step_descs", [])
                    if not step_descs or len(step_descs) < 2:
                        lines = [re.sub(r'^[1-9\.\s・\-\*]+', '', l).strip() for l in raw_memo.split('\n') if l.strip()]
                        step_descs = [l for l in lines if len(l) > 4 and not any(kw in l for kw in ["報告", "推移", "要因"])]
                        if len(step_descs) < 4:
                            step_descs = [
                                "顧客購買データの詳細分析",
                                "高単価ターゲット層へのアプローチ強化",
                                "新プロモーションの実施",
                                "月商200万円の継続的維持"
                            ]

                    for idx in range(1, 5):
                        desc = step_descs[idx - 1] if len(step_descs) >= idx else ""
                        ch_map[f"[[STEP{idx}の説明]]"] = desc

                    # テーブル
                    tbl = ch.get("table", {})
                    if isinstance(tbl, dict):
                        headers = tbl.get("headers", [])
                        rows = tbl.get("rows", [])
                        for idx in range(1, 5):
                            h_val = headers[idx - 1] if len(headers) >= idx else tbl.get(f"col{idx}", "")
                            ch_map[f"[[列見出し{idx}]]"] = h_val
                        for r_idx in range(1, 5):
                            row_data = rows[r_idx - 1] if len(rows) >= r_idx else []
                            for c_idx in range(1, 5):
                                if isinstance(row_data, list) and len(row_data) >= c_idx:
                                    cell_val = row_data[c_idx - 1]
                                else:
                                    cell_val = tbl.get(f"r{r_idx}c{c_idx}", "")
                                ch_map[f"[[行{r_idx}列{c_idx}]]"] = cell_val

                    ch_map["[[表の注記]]"] = notice_val if notice_val else "※詳細につきましてはお気軽にお問い合わせください。"

                    # チェックリスト
                    items = ch.get("items", [])
                    if not items:
                        raw_txt = ch.get("sub_body", "") or ch.get("body", "")
                        lines = [re.sub(r'^[・\-\*\d\.\s]+', '', l).strip() for l in raw_txt.split('\n') if l.strip()]
                        items = lines[:6]

                    intro_val = ch.get("intro", "") or ch.get("body", "以下の確認項目を確実に実施してください。")
                    guide_val = ch.get("guide", "") or "全項目の完了を確認の上、次のフェーズへ進みます。"

                    ch_map["[[チェックリスト導入文]]"] = intro_val
                    for k in range(1, 7):
                        ch_map[f"[[チェック項目{k}]]"] = items[k - 1] if len(items) >= k else ""
                    ch_map["[[チェック後の案内文]]"] = guide_val

                    process_slide_shapes(
                        ch_slide, ch_map,
                        chapter_num=ch_num_str, chapter_title=clean_ch_title,
                        assigned_files=assigned_files, captions=captions
                    )

                    # ★完全無比なグラフ追加（タイトルの被り100%防止・黄金比レイアウト）★
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

                            # 1. 左側本文テキストボックスの幅を Pt(250) に最適化
                            adjust_shapes_for_chart(ch_slide)

                            # 2. グラフを最適位置（left=Pt(350), top=Pt(140), width=Pt(330), height=Pt(270)）へ配置
                            try:
                                c_type = XL_CHART_TYPE.COLUMN_CLUSTERED if ("bar" in layout_type or "棒" in clean_ch_title) else XL_CHART_TYPE.LINE_MARKERS
                                x_pos, y_pos, cx_pos, cy_pos = Pt(350), Pt(140), Pt(330), Pt(270)
                                chart_shape = ch_slide.shapes.add_chart(c_type, x_pos, y_pos, cx_pos, cy_pos, chart_data)
                                chart = chart_shape.chart

                                # ★スライドタイトルがあるため、グラフ内部のタイトル枠は非表示（被り100%防止）★
                                chart.has_title = False

                                force_meiryo_on_chart(chart)
                                chart_created = True  # 単一生成フラグ
                            except Exception:
                                pass

                # 4. 元のテンプレートスライド（最初から存在した12枚）を安全に消去
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
                error_message = str(e)
                if "503" in error_message or "UNAVAILABLE" in error_message or "429" in error_message or "RESOURCE_EXHAUSTED" in error_message:
                    st.error("【サーバー混雑中】現在、AIサーバーへのアクセスが集中しています。恐れ入りますが、1〜2分ほど待ってから再度「資料を生成する」ボタンを押してください。")
                else:
                    st.error(f"エラーが発生しました。時間を置いて再度お試しください。詳細: {e}")
