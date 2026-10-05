import io
import json
import re
import time
import streamlit as st
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.oxml import OxmlElement
from pptx.oxml.xmlchemy import OxmlElement as create_oxml_elem

# ==========================================
# 1. XMLレベルでのフォント強制・クローンユーティリティ
# ==========================================

def apply_chart_meiryo_font(chart):
    """
    DrawingML (a:ea / a:latin) レベルで 'メイリオ' / 'Meiryo' を直接注入し、
    PowerPoint側で標準フォントへ強制初期化される現象を完全に防止する。
    """
    chart_elem = chart._element
    
    def enforce_meiryo(elem):
        if elem is None:
            return
        # 東アジア言語フォント (a:ea)
        ea = elem.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ea')
        if ea is None:
            ea = create_oxml_elem('a:ea')
            elem.append(ea)
        ea.set('typeface', 'メイリオ')
        
        # 欧文フォント (a:latin)
        latin = elem.find('{http://schemas.openxmlformats.org/drawingml/2006/main}latin')
        if latin is None:
            latin = create_oxml_elem('a:latin')
            elem.append(latin)
        latin.set('typeface', 'Meiryo')

    # グラフ内のすべての run properties (rPr) および default run properties (defRPr) に適用
    for rPr in chart_elem.xpath('.//a:rPr'):
        enforce_meiryo(rPr)
    for defRPr in chart_elem.xpath('.//a:defRPr'):
        enforce_meiryo(defRPr)


def duplicate_slide_safe(prs, source_slide_index):
    """
    指定インデックスのスライドを複製し、Shape ID重複および孤立リレーション(rels)を切断した
    修復ダイアログが発生しない安全なスライドオブジェクトを生成する。
    """
    source_slide = prs.slides[source_slide_index]
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    new_slide = prs.slides.add_slide(blank_layout)

    # 背景設定の複製
    if source_slide.background.fill.type:
        new_slide.background.fill.copy(source_slide.background.fill)

    # シェイプの複製
    for shape in source_slide.shapes:
        new_el = shape.element.clone()
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

    return new_slide


# ==========================================
# 2. APIリトライ & 生成データ処理
# ==========================================

def generate_content_with_retry(model, prompt, max_retries=4):
    """
    Gemini APIのレート制限（HTTP 429/503等）に対する指数バックオフ自動再試行処理。
    """
    delay = 3
    for attempt in range(max_retries):
        try:
            response = model.generate_content(prompt)
            return response.text
        except Exception as e:
            err_msg = str(e)
            if ("429" in err_msg or "503" in err_msg or "Quota" in err_msg) and attempt < max_retries - 1:
                st.warning(f"APIレート制限を検知しました。{delay}秒後に自動再試行します... (試行 {attempt + 1}/{max_retries})")
                time.sleep(delay)
                delay *= 2
            else:
                raise e


def parse_json_response(raw_text):
    """
    GeminiからのレスポンスからJSONブロックを抽出し辞書型に変換する。
    """
    cleaned = re.sub(r'```json\s*', '', raw_text)
    cleaned = re.sub(r'```\s*$', '', cleaned).strip()
    match = re.search(r'\{.*\}', cleaned, re.DOTALL)
    if match:
        cleaned = match.group(0)
    return json.loads(cleaned)


# ==========================================
# 3. レイアウト調整・グラフ・STEP更新処理
# ==========================================

def adjust_layout_and_insert_chart(slide, chart_type, chart_data, left=Inches(4.8), top=Inches(1.8), width=Inches(4.5), height=Inches(3.8)):
    """
    文字とグラフの重なりを防止するため、スライド左側のテキストボックス幅を Inches(4.0) に収め、
    右側の領域に指定したグラフを配置・メイリオフォント処理を適用する。
    """
    # 既存のテキスト領域の幅を縮小し被りを回避
    for shape in slide.shapes:
        if shape.has_text_frame:
            # 画面左側に配置されているテキスト領域の幅を抑制
            if shape.left < Inches(4.5) and shape.width > Inches(3.8):
                shape.width = Inches(3.9)

    # グラフの挿入
    chart_shape = slide.shapes.add_chart(chart_type, left, top, width, height, chart_data)
    chart = chart_shape.chart

    # メイリオフォントの完全適用
    apply_chart_meiryo_font(chart)
    return chart


def update_step_slide_content(slide, step_items):
    """
    STEPページ（今後のアクション等）において、タイトル・ステップ番号だけでなく
    各STEPの本文（詳細テキスト）を欠落なく確実に反映させる。
    """
    # テキスト保持シェイプの特定（Y座標・X座標でソート）
    text_shapes = [s for s in slide.shapes if s.has_text_frame]
    text_shapes.sort(key=lambda s: (s.top, s.left))

    step_idx = 0
    for shape in text_shapes:
        tf = shape.text_frame
        text = tf.text.strip()

        # STEPの本文用プレースホルダーまたは本文テキストエリアを判定して更新
        if step_idx < len(step_items):
            item = step_items[step_idx]

            # タイトル枠または本文枠を順次置換
            if "STEP" in text or "ステップ" in text or "詳細" in text or len(text) == 0:
                # 該当枠へタイトルと本文を注入
                tf.clear()
                p1 = tf.paragraphs[0]
                p1.text = item.get("title", f"STEP {step_idx + 1}")
                p1.font.bold = True
                p1.font.name = "メイリオ"
                p1.font.size = Pt(14)
                p1.font.color.rgb = RGBColor(0, 51, 102)

                if "desc" in item and item["desc"]:
                    p2 = tf.add_paragraph()
                    p2.text = item["desc"]
                    p2.font.name = "メイリオ"
                    p2.font.size = Pt(10)
                    p2.font.color.rgb = RGBColor(51, 51, 51)

                step_idx += 1


# ==========================================
# 4. メイン構築ロジック
# ==========================================

def build_presentation(template_path, data, output_path):
    prs = Presentation(template_path)
    
    # グラフの重複挿入を防止するためのフラグ制御
    chart_inserted = False

    slides_data = data.get("slides", [])

    for idx, slide_info in enumerate(slides_data):
        # テンプレート枚数を超える場合は安全複製して拡張
        if idx < len(prs.slides):
            slide = prs.slides[idx]
        else:
            slide = duplicate_slide_safe(prs, len(prs.slides) - 1)

        # 1. タイトル & サブタイトルの更新
        if "title" in slide_info:
            for shape in slide.shapes:
                if shape.has_text_frame and shape.text_frame.text.strip():
                    # 最初の主要テキストボックスをタイトルとみなす
                    shape.text_frame.paragraphs[0].text = slide_info["title"]
                    for p in shape.text_frame.paragraphs:
                        p.font.name = "メイリオ"
                    break

        # 2. 本文テキストの更新
        if "body_bullets" in slide_info and isinstance(slide_info["body_bullets"], list):
            for shape in slide.shapes:
                if shape.has_text_frame and shape != slide.shapes[0]:
                    tf = shape.text_frame
                    tf.clear()
                    for bullet in slide_info["body_bullets"]:
                        p = tf.add_paragraph()
                        p.text = bullet
                        p.font.name = "メイリオ"
                        p.font.size = Pt(12)
                        p.font.color.rgb = RGBColor(51, 51, 51)
                    break

        # 3. STEP（今後のアクション）ページの補完
        if slide_info.get("type") == "step" or "steps" in slide_info:
            steps = slide_info.get("steps", [])
            if steps:
                update_step_slide_content(slide, steps)

        # 4. グラフ挿入の制御（重複防止: 1回のみ生成かつ数値データが存在する対象スライドに限定）
        if not chart_inserted and "chart_data" in slide_info and slide_info["chart_data"]:
            c_info = slide_info["chart_data"]
            categories = c_info.get("categories", [])
            series_list = c_info.get("series", [])

            if categories and series_list:
                chart_data = CategoryChartData()
                chart_data.categories = categories
                for s in series_list:
                    chart_data.add_series(s.get("name", "データ"), tuple(s.get("values", [])))

                chart_type_str = c_info.get("type", "COLUMN").upper()
                c_type = XL_CHART_TYPE.COLUMN_CLUSTERED
                if "LINE" in chart_type_str:
                    c_type = XL_CHART_TYPE.LINE_MARKERS

                # レイアウト自動調整の上でグラフ挿入＆メイリオXML適用
                adjust_layout_and_insert_chart(slide, c_type, chart_data)
                
                # フラグを立てて2度目以降のグラフ挿入をブロック
                chart_inserted = True

    prs.save(output_path)


# ==========================================
# 5. Streamlit UI エントリポイント
# ==========================================

def main():
    st.set_page_config(page_title="PowerPoint自動生成アプリ", layout="wide")
    st.title("📊 PowerPoint 提案書自動生成アプリ")

    st.markdown("""
    提案概要を入力することで、テンプレート構造に基づいたPowerPointプレゼンテーションを自動生成します。
    """)

    # サイドバー: 設定項目
    st.sidebar.header("⚙️ 設定・APIキー")
    api_key = st.sidebar.text_input("Gemini API Key", type="password")
    
    template_file = st.sidebar.file_uploader("テンプレートファイル (.pptx)", type=["pptx"])
    template_path = "資料作成テンプレート_A4横.pptx"

    if template_file:
        with open("uploaded_template.pptx", "wb") as f:
            f.write(template_file.getbuffer())
        template_path = "uploaded_template.pptx"

    # メインフォーム
    prompt_input = st.text_area(
        "提案内容・事業テーマを入力してください",
        height=150,
        placeholder="例: 飲食チェーン向けAI需要予測システムの導入提案。売上推移と成長戦略を含め、導入STEPを明確にする。"
    )

    if st.button("🚀 プレゼンテーション生成", type="primary"):
        if not api_key:
            st.error("Gemini APIキーを入力してください。")
            return
        if not prompt_input.strip():
            st.warning("提案内容を入力してください。")
            return

        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')

        system_prompt = f"""
        あなたはプロフェッショナルな資料作成コンサルタントです。
        入力された提案内容に基づき、全12枚のプレゼンテーション用JSONデータを作成してください。

        【制約事項】
        1. 必ず以下のValidなJSONフォーマットのみを出力してください。
        2. グラフデータ("chart_data")は数値分析が必要な1枚のスライドのみに含め、他のスライドでは null としてください。（重複生成の防止）
        3. STEPスライド("type": "step")では、各ステップの title と desc（本文詳細）を必ず両方出力してください。

        【出力フォーマット構造】
        {{
            "slides": [
                {{
                    "slide_num": 1,
                    "title": "タイトル",
                    "body_bullets": ["ポイント1", "ポイント2"]
                }},
                {{
                    "slide_num": 2,
                    "title": "売上推移と実績分析",
                    "type": "chart",
                    "chart_data": {{
                        "type": "COLUMN",
                        "categories": ["2022", "2023", "2024", "2025"],
                        "series": [
                            {{"name": "売上高(百万円)", "values": [120, 150, 200, 280]}}
                        ]
                    }},
                    "body_bullets": ["過去3年で売上は2.3倍に成長", "AI導入によりさらなる拡大が見込める"]
                }},
                {{
                    "slide_num": 3,
                    "title": "今後の実行プロセス",
                    "type": "step",
                    "steps": [
                        {{"title": "STEP 1: 要件定義", "desc": "現状プロセスのヒアリングと課題抽出を実施"}},
                        {{"title": "STEP 2: モデル構築", "desc": "過去データに基づく予測アルゴリズムの検証"}},
                        {{"title": "STEP 3: 本格運用", "desc": "現場オペレーションへの組み込みと定着化"}},
                        {{"title": "STEP 4: 効果検証", "desc": "KPI達成度評価とモデル改修" viewer}}
                    ]
                }}
            ]
        }}

        提案内容: {prompt_input}
        """

        with st.spinner("AIが提案構造を作成中... (APIリトライ制御適用中)"):
            try:
                raw_json = generate_content_with_retry(model, system_prompt)
                parsed_data = parse_json_response(raw_json)

                output_filename = "generated_presentation.pptx"
                build_presentation(template_path, parsed_data, output_filename)

                st.success("🎉 プレゼンテーションの生成が完了しました！")

                with open(output_filename, "rb") as f:
                    st.download_button(
                        label="📥 完成したPowerPointをダウンロード",
                        data=f,
                        file_name="提案資料_完成版.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )

            except Exception as e:
                st.error(f"エラーが発生しました: {str(e)}")


if __name__ == "__main__":
    main()
