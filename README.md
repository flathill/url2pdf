# セキュリティ監査エビデンス収集・マーキングシステム

## 概要

セキュリティ監査や要件適合性チェックにおいて、要件一覧（Excel）を起点に、Webページの証跡収集からPDFへのマーキング・翻訳付与までを自動化するツールチェーンです。

人間が行うのは「要件の記入」と「AIエージェントへの指示」のみ。残りの工程はすべてスクリプトが自動処理します。

## システム構成

本システムは以下のコンポーネントで構成されています。

| コンポーネント | 種別 | 役割 |
|---|---|---|
| prompt.md | AI指示書 | ステップ2でAIエージェントが参照する指示書 |
| assign_filenames.py | スクリプト | F列URLに通し番号ファイル名を割当、G/H列を置換、J列に書込み |
| url2pdf.py | スクリプト | F列URLをヘッドレスChromiumでPDF化し pdf/ に出力 |
| pdf_annotate.py | スクリプト | PDF内テキスト検索→ハイライト・項番・翻訳を追記、抽出PDFを生成 |

入力はExcelファイル1つ、出力は pdf/（生PDF）と pdf_annotated/（注釈付きPDF＋抽出PDF）です。

## ワークフロー

```
ステップ1  🧑 手作業      人間がA〜D列（大分類・中分類・項番・要件）を記入
    ↓
ステップ2  🤖 AI          AIエージェントがE〜H列を一括記入（prompt.md に従う）
    ↓
ステップ3  ⚙️ 数秒        assign_filenames.py でURL→ファイル名割当、G/H列置換
    ↓
ステップ4  ⚙️ 数分        url2pdf.py でURL→PDF変換
    ↓
ステップ5  ⚙️ 数十秒      pdf_annotate.py でハイライト・項番・翻訳追記＋ページ抽出
```

## Excel 列マップ

| 列 | 内容 | 記入者 | 用途 |
|:---:|---|---|---|
| A | 大分類 | 人間 | 分類整理 |
| B | 中分類 | 人間 | 分類整理 |
| C | 項番 | 人間 | 赤ラベルの表示テキスト |
| D | 要件 | 人間 | 監査要件の本文 |
| E | 回答 | AI | 要件に対する簡潔な回答 |
| F | URL | AI | 証跡元のWebページURL |
| G | 証跡テキスト原文 | AI → スクリプト置換 | `[ファイル名\|原文テキスト]` 形式 |
| H | 日本語訳 | AI → スクリプト置換 | `[ファイル名\|日本語訳]` 形式 |
| I | ページ番号 | — | 現在未使用 |
| J | ファイル名 | スクリプト自動生成 | `assign_filenames.py` が書込み |

各ステップにおけるExcelの具体的な記入状態は [`samples/sample.md`](samples/sample.md) を参照してください。

## リポジトリ構成

```
.
├── README.md                ← 本ファイル（全体説明）
├── prompt.md                ← AIエージェントへの指示書（ステップ2用）
├── assign_filenames.py      ← URL→ファイル名割当スクリプト（ステップ3）
├── assign_filenames.py.md   ← assign_filenames.py の詳細ドキュメント
├── url2pdf.py               ← URL→PDF変換スクリプト（ステップ4）
├── url2pdf.py.md            ← url2pdf.py の詳細ドキュメント
├── pdf_annotate.py          ← PDF注釈・抽出スクリプト（ステップ5）
├── pdf_annotate.py.md       ← pdf_annotate.py の詳細ドキュメント
└── samples/                 ← サンプルファイル
    ├── sample.md            ← サンプルの説明・設計意図
    ├── sample_step1.csv     ← ステップ1完了時点のExcel
    ├── sample_step2.csv     ← ステップ2完了時点のExcel
    └── sample_step3.csv     ← ステップ3完了時点のExcel（完成形）
```

実行時に以下のディレクトリ・ファイルが生成されます。

```
├── pdf/                     ← url2pdf.py が生成する生PDF群
├── pdf_annotated/           ← pdf_annotate.py が生成する注釈付きPDF群
│   ├── EVD001.pdf           ← 全ページ注釈付き
│   ├── EVD001_抽出.pdf       ← コメント箇所±1ページのみ（レビュー用）
│   └── ...
└── .text_index_cache.json   ← テキストインデックスのキャッシュ
```

## 各スクリプトの役割

### assign_filenames.py

F列のURLをスキャンし、URLに含まれるキーワード（例: `zabbix` → `ZBX`、`veeam` → `VEM`）に基づいてプレフィックス付き通し番号ファイル名を割り当てます。G列・H列の `[URL|テキスト]` を `[ファイル名|テキスト]` に自動置換し、J列にファイル名を書き込みます。元ファイルは変更せず、タイムスタンプ付き別名で出力します。

詳細: [`assign_filenames.py.md`](assign_filenames.py.md)

### url2pdf.py

Playwright（ヘッドレスChromium）を使用してURLをPDF化します。Cookieバナーの自動除去、遅延読み込み画像の強制読込、SPA/Docusaurusレイアウト修正、ヘッダー・フッター付与を行います。並列ワーカー（デフォルト5）で高速処理し、既存PDFはスキップします。

詳細: [`url2pdf.py.md`](url2pdf.py.md)

### pdf_annotate.py

生成されたPDF内でG列のテキストをパターンマッチ検索し、3種類のFreeTextアノテーションを追加します。黄色ハイライト（全マッチ箇所）、赤色項番ラベル（ページ内重複排除）、青色日本語訳（常に表示）です。全アノテーションはPDFビューアで編集・移動・削除可能です。処理後、注釈ページ±1ページのみを抽出した軽量レビュー用PDF（`*_抽出.pdf`）も生成します。

詳細: [`pdf_annotate.py.md`](pdf_annotate.py.md)

## 前提条件

Python 3.8以上が必要です。各スクリプトの依存パッケージと環境構築手順の詳細は、それぞれの詳細ドキュメントを参照してください。

```bash
pip install openpyxl playwright pymupdf
playwright install chromium
```

| パッケージ | 用途 | 使用スクリプト |
|---|---|---|
| `openpyxl` | Excel（.xlsx）の読み書き | 全スクリプト |
| `playwright` | ヘッドレスChromiumによるPDF生成 | url2pdf.py |
| `pymupdf` | PDFの読み込み・注釈追記・ページ操作 | pdf_annotate.py |

## クイックスタート

```bash
# 環境構築
python3 -m venv url2pdf-env
source url2pdf-env/bin/activate
pip install openpyxl playwright pymupdf
playwright install chromium

# ステップ1: input.xlsx の A〜D列を記入（人間）
# ステップ2: AIエージェントに prompt.md と共に input.xlsx を渡し、E〜H列を埋める

# ステップ3: ファイル名割当
python assign_filenames.py input_filled.xlsx --dry-run   # 事前確認
python assign_filenames.py input_filled.xlsx              # 実行

# ステップ4: PDF生成（出力ファイル名はステップ3の実行結果を確認）
python url2pdf.py input_filled_named_YYYYMMDD_HHMMSS.xlsx -o ./pdf

# ステップ5: アノテーション＆抽出
python pdf_annotate.py input_filled_named_YYYYMMDD_HHMMSS.xlsx \
    --input-dir=./pdf \
    --output-dir=./pdf_annotated
```

## 再実行時の動作

すべてのスクリプトは冪等性を持ちます。`assign_filenames.py` はタイムスタンプ付き別名で出力するため元ファイルを破壊しません。`url2pdf.py` は既存PDF（60KB以上）をスキップします。`pdf_annotate.py` は出力先に既存ファイルがあればスキップし、抽出PDFも同様です。強制的に再処理したい場合は出力ディレクトリを削除してから再実行してください。

## 既知の制約

G列のテキストはPDF内の実テキストと正確に一致する必要があり、AIによる要約や言い換えが含まれるとマッチに失敗します。`url2pdf.py` はログイン認証が必要なページには対応していません。各スクリプト固有の制約や注意事項は、それぞれの詳細ドキュメントを参照してください。

## 詳細ドキュメント

| ファイル | 内容 |
|---|---|
| [`prompt.md`](prompt.md) | AIエージェントへの指示書（ステップ2用） |
| [`assign_filenames.py.md`](assign_filenames.py.md) | ファイル名割当スクリプトの詳細説明 |
| [`url2pdf.py.md`](url2pdf.py.md) | PDF変換スクリプトの詳細説明 |
| [`pdf_annotate.py.md`](pdf_annotate.py.md) | アノテーションスクリプトの詳細説明 |
| [`samples/sample.md`](samples/sample.md) | サンプルファイルの説明・設計意図 |
