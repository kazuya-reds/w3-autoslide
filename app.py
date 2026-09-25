import streamlit as st
from google import genai
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from PIL import Image
import json
import datetime
import io
import re
import time
import copy

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
    """429(Quota超過)時に指定時間待機して自動リトライする安全関数"""
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
    """注意文の先頭にある『※』『*』『注意：』などの重複記号を自動除去"""
    if not text:
        return ""
    return re.sub(r'^(?:[※\*\:\s]|注意[：:]|注[：:])+', '', str(text).strip())

# ==========================================
# テンプレート・スライド物理複製クローンエンジン
# ==========================================
def duplicate_slide_in_prs(prs, src_idx):
    """
    テンプレートの指定インデックスのスライドを複製（クローン）して
    プレゼンテーションの末尾に安全に追加する関数
    """
    src_slide = prs.slides[src_idx]
    new_slide = prs.slides.add_slide(src_slide.slide_layout)
    
    # 自動追加された不要なデフォルト枠をクリア
    for shp in list(new_slide.shapes):
        sp = shp._element
        sp.getparent().remove(sp)
        
    # ソーススライドの全シェイプ（図形・テキスト・画像・表）を複製
    for shape in src_slide.shapes:
        new_el = copy.deepcopy(shape._element)
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                rId = shape._element.xpath('.//@r:embed')[0]
                if rId in src_slide.part.rels:
                    rel = src_slide.part.rels[rId]
                    new_rId = new_slide.part.relate_to(rel.target_part, rel.reltype)
                    for elem in new_el.iter():
                        for k, v in list(elem.attrib.items()):
                            if k.endswith('embed'):
                                elem.attrib[k] = new_rId
            except Exception:
                pass
        new_slide.shapes._spTree.append(new_el)
    return new_slide

# ==========================================
# ページ番号確実配置エンジン
# ==========================================
def ensure_page_number(slide, page_num):
    """スライド右下にページ番号（1, 2, 3...）を確実に付与"""
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
# 1. 入力UI
# ==========================================
with st.sidebar:
    st.header("⚙️ 設定")
    api_key = st.secrets["GEMINI_API_KEY"]
    st.info("テンプレート: 資料作成テンプレート_A4横.pptx を使用します")

col1, col2 = st.columns(2)
with col1:
    client_name = st.text_input("提案先（顧客名）", "例）役員会 各位")
    doc_title = st.text_input("資料タイトル", "例）経費精算システム導入による業務効率化提案")
with col2:
    sender_name = st.text_input("作成者・部署など", "例）経理部・DX推進チーム")
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
"""例）
経費精算の件、現場から不満多すぎるからどうにかしたい。
スマホで領収書パシャって撮ったら終わるようにできないかな？
今、紙で提出してもらって経理で1枚ずつチェックしてるけど、月末マジで地獄。ミスも多いし。

やりたいこと・アイデア：
・経費精算システムの導入（クラウドのやつ）
・OCR機能で領収書を自動読み取り
・上長承認もスマホでポチッとできるようにしたい

期待できる効果：
経理の確認作業が月15時間くらい減るはず。
社員も外出先から申請できるから「提出遅れ」が無くなる。

費用は月数万円くらいなら出せる？
来月にはツール決めて、再来月から一部部署でテスト運用したい感じ。
予算感とスケジュールまとめた提案スライド作って上に通したい。""", height=400)


# ==========================================
# 2. 段落処理（テンプレート書式完全維持）
# ==========================================
def process_paragraph_runs(p, replace_map, chapter_num=None, chapter_title=None):
    clean_title = chapter_title
    if clean_title:
        clean_title = re.sub(r'^\d+[\.\s_]*', '', str(clean_title)).strip()

    for r in p.runs:
        if chapter_num and "00" in r.text:
            r.text = r.text.replace("00", chapter_num)
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
            full_text = full_text.replace("[[章タイトル]]", clean_title)
        for tag, val in replace_map.items():
            if tag in full_text:
                full_text = full_text.replace(tag, str(val))

        if len(p.runs) > 0:
            p.runs[0].text = full_text
            for r in p.runs[1:]:
                r.text = ""


# ==========================================
# 3. スライド要素処理（動的・画像・コネクタ線精密判別）
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

        # 1. 注意文・補足文がメモに無い場合のみ、注意アイコン画像や注意専用テキスト枠を削除
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

        # 2. 補助項目（sub_title / sub_body）にデータが無い場合のみ、補助項目枠を削除
        # ※ 補助項目にデータがある場合は、縦線も含めて絶対に削除しない（保護維持）
        if not has_sub_item and shape.has_text_frame:
            if "[[補助項目]]" in raw_text or "[[補助項目の本文]]" in raw_text:
                shapes_to_remove.append(shape)
                continue

        # 3. 本文がメモに無い場合の消去
        if not has_body and shape.has_text_frame:
            if "[[本文]]" in raw_text:
                shapes_to_remove.append(shape)
                continue

        # 4. 画像枠処理（アップロードされた画像のみ差し替え）
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

        # 5. テキストボックス処理
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

        # 6. 表（テーブル）処理
        if shape.has_table:
            for cell in shape.table.iter_cells():
                for p in cell.text_frame.paragraphs:
                    process_paragraph_runs(p, replace_map)

    # 不要図形の物理除去
    for shp in shapes_to_remove:
        try:
            sp = shp._element
            sp.getparent().remove(sp)
        except Exception:
            pass


# ==========================================
# 4. メイン処理（完全汎用クローン生成）
# ==========================================
if st.button("✨資料を生成する✨"):
    if not api_key:
        st.error("APIキーを入力してください。")
    elif not raw_memo.strip():
        st.error("提案メモを入力してください。")
    else:
        status_box = st.empty()
        with st.spinner("AIがメモを解析し、資料を構築中..."):
            try:
                client = genai.Client(api_key=api_key)

                prompt_data = {
                    "client_name": client_name,
                    "doc_title": doc_title,
                    "raw_memo": raw_memo,
                    "image_contexts": image_contexts
                }

                # 高精度構造解析プロンプト（完全ドメインフリー）
                single_pass_prompt = f"""
                あなたはプレゼン資料作成のプロコンサルタントです。
                以下の入力を直接分析し、メモの情報を一切落とさずにプレゼンスライド用のJSONを作成してください。

                【入力データ】
                {json.dumps(prompt_data, ensure_ascii=False, indent=2)}

                【レイアウト＆情報抽出ルール】
                1. メモの主要項目を整理し、プレゼンに最適な3〜7個の章（chapters）にまとめてください。
                2. 各章のコンテンツに合わせて、最適な "layout_type" を以下から選定してください:
                   - "text": 概要・課題・特徴・メリット・期待効果など
                   - "table": 料金プラン・比較表・一覧データなど
                   - "step": 導入手順・運用フロー・タイムラインなど（最大4ステップ）
                   - "checklist": 確認事項・必要書類・チェックリストなど
                3. 情報の分配ルール:
                   - "main_item": キャッチコピーまたは主要項目（15文字以内）
                   - "body": メインの概要文章
                   - "sub_title": 補助項目のタイトル（例: "主な課題点", "提供する主なコンテンツ", "対象別の具体的な効果", "今後の展開領域" など）。メモに箇条書きや複数カテゴリがある場合は必ずここにタイトルを入れてください。
                   - "sub_body": 補助項目の本文（例: 箇条書きリスト、店舗側・新人側の詳細効果など）
                   - "notice": 注意事項（※記号の文など）。無ければ空文字 ""
                4. "step" の場合:
                   - "step_descs": [ "ステップ1の説明", "ステップ2の説明", "ステップ3の説明", "ステップ4の説明" ] の配列形式で格納してください。
                5. "table" の場合:
                   - "column_count": 2, 3, 4 のいずれか
                   - "table": {{ "headers": ["列見出し1", "列見出し2", ...], "rows": [ ["行1列1", "行1列2", ...], ["行2列1", ...] ] }} の形式で格納してください。
                6. 【重要ルール】
                   - "body" や "sub_body" などの本文項目は絶対に空欄にせず、文字を出力してください。
                   - 「提案内容」などのスライドでは、入力データのメモにある具体的なアイデア（OCR機能など）を漏れなく抽出し、必ず文章や箇条書きで反映させてください。

                【出力形式】
                純粋なJSONのみを出力してください。
                {{
                  "chapters": [
                    {{
                      "chapter_title": "章タイトル",
                      "layout_type": "text",
                      "main_item": "主要見出し（15文字以内）",
                      "body": "本文テキスト（※絶対に空欄にせず、入力データから抽出した内容を必ず2〜4点の箇条書き等で具体的に記述すること。情報が足りない場合はビジネスの文脈から推測してでも必ず出力すること）",
                      "sub_title": "補助項目タイトル",
                      "sub_body": "補助項目本文",
                      "notice": "",
                      "step_descs": [],
                      "table": {{ "headers": [], "rows": [] }},
                      "assigned_image_indices": [],
                      "image_captions": []
                    }}
                  ]
                }}
                """

                txt_res = call_gemini_with_retry(client, single_pass_prompt)
                m_json = re.search(r'\{[\s\S]*\}', txt_res)
                final_data = json.loads(m_json.group(0) if m_json else txt_res)

                status_box.success("✅ AI構造解析完了：PowerPointスライドを複製・構築中...")

                prs = Presentation('資料作成テンプレート_A4横.pptx')
                tpl_slide_count = len(prs.slides)
                
                chapters = final_data.get("chapters", [])
                today_str = datetime.date.today().strftime("%Y/%m/%d")

                # 生成スライドの並び順管理
                new_slides = []

                # --- 1. 表紙スライドの複製・生成 (Template Index 0) ---
                cover_slide = duplicate_slide_in_prs(prs, 0)
                cover_map = {
                    "[[資料タイトル]]": doc_title,
                    "[[顧客名]]": client_name,
                    "[[バージョン]]": "1.0",
                    "[[更新日]]": today_str
                }
                process_slide_shapes(cover_slide, cover_map)
                new_slides.append(cover_slide)

                # --- 2. INDEX（目次）スライドの複製・生成 (★2ページ目に配置) ---
                index_slide = duplicate_slide_in_prs(prs, 1)
                new_slides.append(index_slide)

                # --- 3. 各章スライドの選定・複製・データ流し込み ---
                index_map = {}
                chapter_slides = []

                for ch_idx, ch in enumerate(chapters):
                    ch_num_str = f"{ch_idx + 1:02d}"  # 01, 02, 03...
                    ch_page_num = ch_idx + 2          # INDEX=1 のため、第1章は P2 スタート
                    ch_title = ch.get("chapter_title", f"章{ch_idx+1}")
                    clean_ch_title = re.sub(r'^\d+[\.\s_]*', '', str(ch_title)).strip()

                    index_map[f"[[章{ch_idx+1}タイトル]]"] = f"{ch_num_str} {clean_ch_title}"
                    index_map[f"P[[章{ch_idx+1}ページ]]"] = f"P{ch_page_num}"
                    index_map[f"[[章{ch_idx+1}ページ]]"] = str(ch_page_num)

                    layout_type = str(ch.get("layout_type", "text")).lower()
                    img_indices = ch.get("assigned_image_indices", [])
                    img_count = len(img_indices)
                    
                    assigned_files = [uploaded_files[i] for i in img_indices if uploaded_files and i < len(uploaded_files)]
                    captions = ch.get("image_captions", [])

                    # テンプレート上の複製元インデックス判定
                    target_template_idx = 2
                    if layout_type == "checklist" or "確認" in clean_ch_title or "チェック" in clean_ch_title:
                        target_template_idx = 9 if tpl_slide_count > 9 else 5
                    elif layout_type == "step" or "フロー" in clean_ch_title or "流れ" in clean_ch_title or "手順" in clean_ch_title:
                        target_template_idx = 5 if tpl_slide_count > 5 else 4
                    elif layout_type == "table" or "プラン" in clean_ch_title or "比較" in clean_ch_title or "料金" in clean_ch_title:
                        tbl_data = ch.get("table", {})
                        cols = ch.get("column_count", len(tbl_data.get("headers", [])))
                        if cols == 4 and tpl_slide_count > 8: target_template_idx = 8
                        elif cols == 3 and tpl_slide_count > 7: target_template_idx = 7
                        elif tpl_slide_count > 6: target_template_idx = 6
                        else: target_template_idx = 3
                    else:
                        if img_count >= 2 and tpl_slide_count > 3: target_template_idx = 3
                        elif img_count == 1 and tpl_slide_count > 4: target_template_idx = 4
                        else: target_template_idx = 2

                    ch_slide = duplicate_slide_in_prs(prs, target_template_idx)
                    chapter_slides.append(ch_slide)

                    # 置換辞書の組み立て
                    ch_map = {}
                    ch_map["[[主要項目]]"] = ch.get("main_item", "")
                    ch_map["[[本文]]"] = ch.get("body", "")
                    ch_map["[[補助項目]]"] = ch.get("sub_title", "")
                    ch_map["[[補助項目の本文]]"] = ch.get("sub_body", "")

                    notice_val = clean_notice_text(ch.get("notice", ""))
                    ch_map["[[補足文]]"] = notice_val
                    ch_map["[[注意文]]"] = notice_val

                    # STEPスライド（1〜4）
                    step_descs = ch.get("step_descs", [])
                    for idx in range(1, 5):
                        desc = step_descs[idx - 1] if len(step_descs) >= idx else ""
                        ch_map[f"[[STEP{idx}の説明]]"] = desc

                    # テーブルスライド
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
                    items = ch.get("items", ["", "", "", "", "", ""])
                    ch_map["[[チェックリスト導入文]]"] = ch.get("intro", "")
                    for k in range(1, 7):
                        ch_map[f"[[チェック項目{k}]]"] = items[k - 1] if len(items) >= k else ""
                    ch_map["[[チェック後の案内文]]"] = ch.get("guide", "")

                    process_slide_shapes(
                        ch_slide, ch_map,
                        chapter_num=ch_num_str, chapter_title=clean_ch_title,
                        assigned_files=assigned_files, captions=captions
                    )

                # --- 4. INDEX（目次）スライドへデータ反映 ---
                for k in range(len(chapters) + 1, 7):
                    index_map[f"[[章{k}タイトル]]"] = ""
                    index_map[f"P[[章{k}ページ]]"] = ""
                    index_map[f"[[章{k}ページ]]"] = ""

                process_slide_shapes(index_slide, index_map)
                for shp in index_slide.shapes:
                    if shp.has_text_frame:
                        for p in shp.text_frame.paragraphs:
                            if p.text.strip() in ["P", "P.", "P-"]:
                                p.text = ""

                # 並び順の確定: [表紙] -> [INDEX] -> [章スライド群...]
                new_slides.extend(chapter_slides)

                # --- 5. 元のテンプレート用スライド（0〜tpl_slide_count-1）の物理削除 ---
                sldIdLst = prs.slides._sldIdLst
                for i in range(tpl_slide_count - 1, -1, -1):
                    elem = sldIdLst[i]
                    prs.part.drop_rel(elem.rId)
                    sldIdLst.remove(elem)

                # --- 6. 残存タグ消去 ＆ ページ番号確実付与 ---
                for s_i, sld in enumerate(prs.slides):
                    for shp in sld.shapes:
                        if shp.has_text_frame:
                            for p in shp.text_frame.paragraphs:
                                if "[[" in p.text and "]]" in p.text:
                                    p.text = re.sub(r'\[\[.*?\]\]', '', p.text)
                    
                    if s_i > 0:
                        ensure_page_number(sld, s_i)  # 表紙=0、INDEX=1, 第1章=2...

                # --- 7. ファイル名の自動生成（資料タイトル反映） ---
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
