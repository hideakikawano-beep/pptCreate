# PowerPoint テンプレート設計ドキュメント

このリポジトリの PowerPoint 自動生成における「テンプレート」の扱い方をまとめたドキュメントです。

## 1. 基本方針：コードがテンプレート

このプロジェクトでは `.pptx` のマスタースライドやテーマファイルを使う**ファイルベースのテンプレートは使用していません**。
代わりに、`python-pptx` を使って **Python コード上でデザイン仕様（配色・フォント・レイアウト）をハードコードする方式**を採用しています。

つまり「テンプレート」＝ `generate_ppt.py` 内の **定数群と描画関数** です。
すべてのスライドは空白レイアウト（`slide_layouts[6]`）の上に、コードで図形・テキストボックスを座標指定して描画されます。

```python
slide = prs.slides.add_slide(prs.slide_layouts[6])  # 完全な白紙レイアウト
_set_bg_white(slide)                                # 背景を白に
```

この方式の利点：
- PowerPoint の既製テーマに依存せず、デザインの一貫性を完全にコードで保証できる
- JSON データさえあれば誰の環境でも同一の見た目が再現される
- バージョン管理（git）でデザイン変更を追跡できる

## 2. デザイン仕様（テンプレートの中身）

### 配色（`generate_ppt.py` 上部のカラー定数）

| 用途 | 色 | 定数 |
|------|-----|------|
| メインカラー（タイトル・見出し・アクセント） | `#2180FF` | `BLUE` |
| 本文テキスト | `#262626` | `SOFT_BLACK` |
| 背景 | `#FFFFFF` | `WHITE` |
| カードのボーダー | `#E0E0E0` | `LIGHT_GRAY` |
| カードの塗り（グレー） | `#F5F5F5` | `FILL_GRAY` |
| カードの塗り（ブルー系） | `#EBF5FF` | `FILL_BLUE_LIGHT` |
| セパレーター線 | `#D0D0D0` | `SEPARATOR_CLR` |

### フォント
- **Noto Sans**（全テキスト共通、`_add_text_box()` 内で固定指定）

### スライド寸法
- 16:9 ワイド画面：`13.33 × 7.5 インチ`
- 外周マージン：`0.6 インチ`

### スライド共通構造
すべてのコンテンツスライドは以下の固定構造を持ちます（座標は定数で管理）：

```
┌─────────────────────────────────────────┐
│ スライドタイトル  (16pt / Blue / Normal)   │ ← Header
│ ──────── セパレーター線 ────────           │
│ キーメッセージ    (20pt / SoftBlack / Bold)│
│                                          │
│ ┌──────────┐  ┌──────────┐               │ ← Body
│ │  カード   │  │  カード   │               │   (Cards)
│ └──────────┘  └──────────┘               │
└─────────────────────────────────────────┘
```

### カードのスタイル
`_draw_card()` で 2 種類を切替：
- `filled`：グレー塗り（`#F5F5F5`）、ボーダーなし
- `bordered`：白塗り＋ライトグレーのボーダー（`#E0E0E0`）

カード内テキスト：
- 見出し：18pt / Bold / Blue
- 本文：12pt / Normal / SoftBlack

## 3. レイアウト自動分岐（テンプレートの肝）

`_draw_body()` が **カード（body_items）の件数に応じてレイアウトを自動的に切り替えます**。
ここが本テンプレートの最大の特徴です。

| カード件数 | レイアウト | カードスタイル |
|-----------|-----------|--------------|
| 3 件以下 | 横一列（Horizontal Row） | `filled` |
| 4 件（必須） | 2×2 グリッド | `bordered` |
| 5 件以上 | 2 列の動的グリッド（Dynamic） | `filled` |

データ作成者はレイアウトを意識する必要がなく、カードの個数を決めるだけで適切な配置になります。

## 4. スライドの種類

| `type` | ビルダー関数 | 内容 |
|--------|------------|------|
| `title` | `_build_title_slide()` | 表紙。左端に Blue アクセントバー、32pt タイトル、サブタイトル |
| `content`（既定） | `_build_content_slide()` | Header + Body の通常スライド |

## 5. データ形式（JSON スキーマ）

スライドの中身は JSON で記述します。テンプレート（コード）が、この JSON を見た目に変換します。

```json
{
  "title": "プレゼンタイトル",
  "subtitle": "サブタイトル / 日付など",
  "slides": [
    {
      "type": "title",
      "title": "メインタイトル",
      "content": "サブタイトルや概要"
    },
    {
      "type": "content",
      "slide_title": "スライドタイトル（16pt Blue）",
      "key_message": "このスライドの結論（20pt Bold、1〜2行）",
      "body_items": [
        {"heading": "カード見出し", "body": "カード本文（統計・具体例含む）"},
        {"heading": "カード見出し", "body": "カード本文"}
      ],
      "notes": "発表者メモ（スピーカーノートに反映）"
    }
  ]
}
```

各フィールドの役割：
- `slide_title` … ヘッダーのタイトル
- `key_message` … ヘッダーの結論メッセージ
- `body_items[].heading` / `body` … 各カードの見出しと本文
- `notes` … スピーカーノート（任意）
- `bullets` … `body_items` がない場合の互換用。箇条書きが各カード本文に変換される

## 6. 生成方法（3つの入口）

テンプレートを使って PPTX を作る方法は 3 通りあります。

### A. 文字起こしから AI で自動生成 — `generate_ppt.py`
文字起こしテキストを Claude API（`claude-opus-4-6`）に渡し、Web 検索でリサーチを補完した上で、
上記スキーマの JSON 構成（8〜12 スライド）を生成し、そのまま PPTX 化します。

```bash
export ANTHROPIC_API_KEY=...    # 必須
python generate_ppt.py sample_transcription.txt
# → sample_transcription_structure.json  (構成JSON)
# → sample_transcription_presentation.pptx (PowerPoint)
```

`SYSTEM_PROMPT` に上記デザインルール（件数→レイアウト対応など）が明記されており、
AI がテンプレート規約に沿った JSON を返すよう制御しています。

### B. 既存の JSON 構成から生成 — `build_from_json.py`
手元の `*_structure.json` から PPTX のみを再生成します（AI 呼び出しなし）。

```bash
python build_from_json.py cross_marketing_notta_structure.json
# → cross_marketing_notta_presentation.pptx
```

### C. Python コードで直接 JSON を組み立て — `build_aws_summit_pptx.py`
スライド内容を Python の dict として直書きし、`create_pptx()` で生成する例。
AI を使わず、人手で構成を完全にコントロールしたい場合に使います。

```bash
python build_aws_summit_pptx.py
# → aws_summit_2026_structure.json + aws_summit_2026_presentation.pptx
```

いずれの入口も、最終的には `generate_ppt.py` の `create_pptx()` を共通で呼び出すため、
**出力されるデザインは常に同一テンプレートに従います**。

## 7. 依存関係

```
anthropic>=0.40.0      # Claude API（入口Aのみ必要）
python-pptx>=1.0.0     # PPTX 生成（全入口で必要）
```

## 8. テンプレートをカスタマイズするには

| 変更したいもの | 編集箇所（`generate_ppt.py`） |
|---------------|------------------------------|
| 配色 | 冒頭の「カラー定数」（`BLUE`, `SOFT_BLACK` など） |
| フォント | `_add_text_box()` の `run.font.name = "Noto Sans"` |
| 余白・各領域の位置 | 「スライド寸法」セクションの定数（`MARGIN`, `TITLE_TOP` など） |
| カードの見た目 | `_draw_card()` |
| 件数→レイアウトのルール | `_draw_body()` |
| 表紙のデザイン | `_build_title_slide()` |
| AI への指示（生成ルール） | `SYSTEM_PROMPT` |
