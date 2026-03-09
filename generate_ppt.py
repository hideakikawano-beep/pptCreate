#!/usr/bin/env python3
"""
文字起こしデータからPowerPointを自動生成するスクリプト

使い方:
    python generate_ppt.py [transcription_file]

例:
    python generate_ppt.py sample_transcription.txt

環境変数:
    ANTHROPIC_API_KEY: Claude API キー（必須）
"""

import sys
import json
import os
import re
from pathlib import Path

import anthropic
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


# ─── スライドスタイル設定 ─────────────────────────────────────────────────────

THEME = {
    "title_bg": RGBColor(0x1A, 0x1A, 0x2E),      # 濃紺
    "section_bg": RGBColor(0x16, 0x21, 0x3E),     # 少し薄い濃紺
    "content_bg": RGBColor(0xFF, 0xFF, 0xFF),      # 白
    "accent": RGBColor(0x0F, 0x3D, 0xAF),          # 青
    "title_text": RGBColor(0xFF, 0xFF, 0xFF),       # 白
    "body_text": RGBColor(0x1A, 0x1A, 0x1A),       # ほぼ黒
    "bullet_accent": RGBColor(0x0F, 0x3D, 0xAF),   # 青（箇条書きマーカー）
}


# ─── 文字起こし読み込み ───────────────────────────────────────────────────────

def read_transcription(file_path: str) -> str:
    with open(file_path, "r", encoding="utf-8") as f:
        return f.read()


# ─── Claude APIでリサーチ＆構成作成 ──────────────────────────────────────────

def research_and_structure(transcription: str) -> dict:
    """
    Claude + Web検索で文字起こしを分析し、プレゼン構成をJSONで返す。
    """
    client = anthropic.Anthropic()

    prompt = f"""あなたは優秀なプレゼンテーションデザイナーです。
以下の会議/講演の文字起こしデータを分析し、ウェブ検索で関連情報を補完して、
プロフェッショナルなプレゼンテーション資料を作成してください。

## 文字起こしデータ:
{transcription}

## 手順:
1. 文字起こしの主要トピック・課題・提言を抽出する
2. 不足している統計データ、最新情報、業界トレンドをウェブ検索で補完する
3. 論理的な流れになるようスライド構成を設計する

## 出力形式（必ずこのJSON形式のみで返すこと）:
{{
  "title": "プレゼンテーションのメインタイトル",
  "subtitle": "サブタイトルや日付・主催者など",
  "slides": [
    {{
      "type": "title",
      "title": "タイトルスライドの見出し",
      "content": "サブタイトルや概要文"
    }},
    {{
      "type": "agenda",
      "title": "アジェンダ",
      "bullets": ["セクション1", "セクション2", "セクション3"]
    }},
    {{
      "type": "section",
      "title": "セクション区切りのタイトル"
    }},
    {{
      "type": "content",
      "title": "スライドタイトル",
      "bullets": [
        "主要ポイント1（具体的な数値や事実を含める）",
        "主要ポイント2",
        "主要ポイント3"
      ],
      "notes": "発表者用メモ（補足情報、調査結果など）"
    }},
    {{
      "type": "summary",
      "title": "まとめ・次のアクション",
      "bullets": ["アクション1", "アクション2", "アクション3"]
    }}
  ]
}}

注意事項:
- スライド枚数は8〜15枚程度
- 各スライドのbulletsは3〜5項目
- ウェブ検索で得た最新の統計・データを積極的に活用する
- 日本語で出力する
- JSON以外のテキスト（説明文、マークダウンのコードブロック記法など）は一切含めないこと"""

    messages = [{"role": "user", "content": prompt}]

    print("Claude が文字起こしを分析・リサーチ中...", flush=True)

    # Web検索ツール付きでストリーミング実行
    with client.messages.stream(
        model="claude-opus-4-6",
        max_tokens=8000,
        thinking={"type": "adaptive"},
        tools=[
            {"type": "web_search_20260209", "name": "web_search"},
        ],
        messages=messages,
    ) as stream:
        current_block_type = None
        for event in stream:
            if not hasattr(event, "type"):
                continue
            if event.type == "content_block_start" and hasattr(event, "content_block"):
                current_block_type = event.content_block.type
                if current_block_type == "thinking":
                    print("  [思考中...]", flush=True)
                elif current_block_type == "tool_use":
                    tool_name = getattr(event.content_block, "name", "")
                    if tool_name == "web_search":
                        print("  [ウェブ検索中...]", flush=True)
                elif current_block_type == "text":
                    print("  [構成を生成中", end="", flush=True)
            elif event.type == "content_block_delta" and hasattr(event, "delta"):
                if event.delta.type == "text_delta":
                    print(".", end="", flush=True)
            elif event.type == "content_block_stop":
                if current_block_type == "text":
                    print("]", flush=True)
                current_block_type = None

        final_message = stream.get_final_message()

    # テキストブロックを結合
    text_content = ""
    for block in final_message.content:
        if block.type == "text":
            text_content += block.text

    # JSON を抽出・パース
    try:
        # コードブロックが混入した場合でも対処
        json_match = re.search(r"\{[\s\S]*\}", text_content)
        if json_match:
            return json.loads(json_match.group())
        return json.loads(text_content)
    except json.JSONDecodeError as e:
        print(f"\nJSON パースエラー: {e}")
        print("レスポンス先頭500文字:", text_content[:500])
        raise


# ─── PowerPoint 生成 ──────────────────────────────────────────────────────────

def _set_bg_color(slide, color: RGBColor):
    """スライドの背景色を設定する。"""
    from pptx.oxml.ns import qn
    from lxml import etree

    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_text_box(slide, text: str, left, top, width, height,
                  font_size: int, bold: bool = False,
                  color: RGBColor = None, align=PP_ALIGN.LEFT):
    """テキストボックスを追加する。"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    return txBox


def _build_title_slide(prs: Presentation, data: dict):
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)
    _set_bg_color(slide, THEME["title_bg"])

    W = prs.slide_width
    H = prs.slide_height

    # アクセントライン
    from pptx.util import Emu
    line = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(0.6), H // 2 - Inches(0.05),
        Inches(1.5), Inches(0.08),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = THEME["accent"]
    line.line.fill.background()

    # メインタイトル
    _add_text_box(
        slide, data.get("title", "プレゼンテーション"),
        Inches(0.6), H // 2 - Inches(1.6),
        W - Inches(1.2), Inches(1.4),
        font_size=36, bold=True,
        color=THEME["title_text"], align=PP_ALIGN.LEFT,
    )

    # サブタイトル
    subtitle = data.get("subtitle", "")
    if subtitle:
        _add_text_box(
            slide, subtitle,
            Inches(0.6), H // 2 + Inches(0.2),
            W - Inches(1.2), Inches(0.8),
            font_size=18,
            color=RGBColor(0xCC, 0xCC, 0xCC), align=PP_ALIGN.LEFT,
        )


def _build_section_slide(prs: Presentation, data: dict):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    _set_bg_color(slide, THEME["section_bg"])

    W = prs.slide_width
    H = prs.slide_height

    # 区切り線
    line = slide.shapes.add_shape(
        1,
        Inches(0.5), H // 2 - Inches(0.6),
        Inches(0.08), Inches(1.2),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = THEME["accent"]
    line.line.fill.background()

    _add_text_box(
        slide, data.get("title", ""),
        Inches(1.0), H // 2 - Inches(0.5),
        W - Inches(1.5), Inches(1.0),
        font_size=32, bold=True,
        color=THEME["title_text"], align=PP_ALIGN.LEFT,
    )


def _build_content_slide(prs: Presentation, data: dict):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    _set_bg_color(slide, THEME["content_bg"])

    W = prs.slide_width
    H = prs.slide_height

    # ヘッダーバー
    header = slide.shapes.add_shape(
        1,
        Inches(0), Inches(0),
        W, Inches(1.2),
    )
    header.fill.solid()
    header.fill.fore_color.rgb = THEME["title_bg"]
    header.line.fill.background()

    # スライドタイトル
    _add_text_box(
        slide, data.get("title", ""),
        Inches(0.4), Inches(0.2),
        W - Inches(0.8), Inches(0.8),
        font_size=24, bold=True,
        color=THEME["title_text"], align=PP_ALIGN.LEFT,
    )

    # アクセントライン（ヘッダー下）
    accent = slide.shapes.add_shape(
        1,
        Inches(0), Inches(1.2),
        W, Inches(0.06),
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = THEME["accent"]
    accent.line.fill.background()

    # 箇条書き
    bullets = data.get("bullets", [])
    if bullets:
        from pptx.util import Pt as PtUtil
        from pptx.oxml.ns import qn
        from lxml import etree

        txBox = slide.shapes.add_textbox(
            Inches(0.5), Inches(1.5),
            W - Inches(1.0), H - Inches(2.0),
        )
        tf = txBox.text_frame
        tf.word_wrap = True

        for i, bullet_text in enumerate(bullets):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()

            p.space_before = Pt(6)
            p.space_after = Pt(6)

            # ビュレット記号
            run_marker = p.add_run()
            run_marker.text = "▶  "
            run_marker.font.size = Pt(14)
            run_marker.font.color.rgb = THEME["bullet_accent"]
            run_marker.font.bold = True

            # テキスト
            run_text = p.add_run()
            run_text.text = bullet_text
            run_text.font.size = Pt(16)
            run_text.font.color.rgb = THEME["body_text"]

    # 発表者メモ
    notes_text = data.get("notes", "")
    if notes_text:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = notes_text


def _build_summary_slide(prs: Presentation, data: dict):
    """まとめスライド（アクセントカラー背景）"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    _set_bg_color(slide, THEME["accent"])

    W = prs.slide_width
    H = prs.slide_height

    _add_text_box(
        slide, data.get("title", "まとめ"),
        Inches(0.5), Inches(0.3),
        W - Inches(1.0), Inches(0.9),
        font_size=28, bold=True,
        color=THEME["title_text"], align=PP_ALIGN.LEFT,
    )

    # 区切り線
    line = slide.shapes.add_shape(
        1,
        Inches(0.5), Inches(1.25),
        W - Inches(1.0), Inches(0.05),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    line.line.fill.background()

    bullets = data.get("bullets", [])
    if bullets:
        txBox = slide.shapes.add_textbox(
            Inches(0.5), Inches(1.5),
            W - Inches(1.0), H - Inches(2.0),
        )
        tf = txBox.text_frame
        tf.word_wrap = True

        for i, bullet_text in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_before = Pt(8)

            run_num = p.add_run()
            run_num.text = f"{i + 1}.  "
            run_num.font.size = Pt(16)
            run_num.font.bold = True
            run_num.font.color.rgb = RGBColor(0xFF, 0xFF, 0xAA)

            run_text = p.add_run()
            run_text.text = bullet_text
            run_text.font.size = Pt(16)
            run_text.font.color.rgb = THEME["title_text"]


def create_pptx(presentation_data: dict, output_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    builders = {
        "title": _build_title_slide,
        "agenda": _build_content_slide,
        "section": _build_section_slide,
        "content": _build_content_slide,
        "summary": _build_summary_slide,
    }

    slides = presentation_data.get("slides", [])
    for slide_data in slides:
        slide_type = slide_data.get("type", "content")
        builder = builders.get(slide_type, _build_content_slide)
        builder(prs, slide_data)

    prs.save(output_path)
    print(f"PowerPoint を保存しました: {output_path}")


# ─── エントリポイント ─────────────────────────────────────────────────────────

def main():
    if len(sys.argv) > 1:
        transcription_file = sys.argv[1]
    else:
        transcription_file = "sample_transcription.txt"

    if not os.path.exists(transcription_file):
        print(f"エラー: ファイルが見つかりません: {transcription_file}")
        sys.exit(1)

    print(f"文字起こしファイル: {transcription_file}")
    transcription = read_transcription(transcription_file)

    # リサーチ＆構成作成
    presentation_data = research_and_structure(transcription)

    # 構成をJSONとして保存（確認用）
    stem = Path(transcription_file).stem
    json_output = f"{stem}_structure.json"
    with open(json_output, "w", encoding="utf-8") as f:
        json.dump(presentation_data, f, ensure_ascii=False, indent=2)
    print(f"プレゼン構成を保存しました: {json_output}")

    # PowerPoint 生成
    pptx_output = f"{stem}_presentation.pptx"
    create_pptx(presentation_data, pptx_output)

    print(f"\n完了！生成ファイル:")
    print(f"  {json_output}  ← スライド構成（JSON）")
    print(f"  {pptx_output}  ← PowerPoint ファイル")


if __name__ == "__main__":
    main()
