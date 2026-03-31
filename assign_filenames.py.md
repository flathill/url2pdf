# assign\_filenames.py

## 概要

セキュリティ監査や要件適合性チェックにおいて、Webページを証跡（エビデンス）としてPDF保存する作業を自動化するワークフローの一部です。

このスクリプトは、Excelファイル内のURL列を全行スキャンし、各URLに対して**重複を排除しながら通し番号付きのファイル名**を自動的に割り当てます。同時に、証跡テキスト列に記載されたURLをファイル名に置換します。

## 背景と目的

監査用Excelでは、1つのURLが複数の要件行から参照されることがよくあります。たとえば、Zabbixのアクセス制御に関するドキュメントページは、「認証」「権限管理」「監査ログ」など複数の要件の証跡として共通的に使われます。

このとき問題になるのが**ファイル名の一意性**です。

- 同じURLに対して異なるファイル名が付くと、同一ページから複数のPDFが生成されて無駄になる
- 人手でファイル名を管理すると、行数が増えたときにミスや不整合が起きる
- AIエージェントが証跡テキストを記入する段階では、Excel全体を俯瞰した重複判定ができない

そこで、**Excel全体を一括スキャンして、URLとファイル名の対応を機械的に確定する**のがこのスクリプトの役割です。

## ワークフロー上の位置づけ

```
ステップ1: 人間がA〜D列（要件定義）を記入
    ↓
ステップ2: AIエージェントがE, F, G, H列を一括で埋める
           F列: 証跡URL
           G列: [URL|英語テキスト]   ← この時点ではファイル名未定
           H列: [URL|日本語訳]       ← 同上
    ↓
★ステップ3: python assign_filenames.py input.xlsx   ← 本スクリプト
           → F列のURLに通し番号ファイル名を割り当て
           → I列にファイル名を書き込み
           → G列・H列の [URL|テキスト] を [ファイル名|テキスト] に置換
    ↓
ステップ4: url2pdf でURL→PDF生成（I列のファイル名を使用）
    ↓
ステップ5: pdf_annotate でPDFにハイライト・翻訳を追記（G/H列を使用）
           → J列に索引情報を出力
```

## Excel列構成

| 列 | 内容 | 本スクリプトでの扱い |
|---|---|---|
| A | 大分類 | 参照しない |
| B | 中分類 | 参照しない |
| C | 項番 | 参照しない |
| D | 要件内容 | 参照しない |
| E | 回答 | 参照しない |
| **F** | **証跡URL** | **読み取り（URLを抽出）** |
| **G** | **英語証跡テキスト** | **読み書き（URLをファイル名に置換）** |
| **H** | **日本語訳** | **読み書き（URLをファイル名に置換）** |
| **I** | **ファイル名** | **書き込み（通し番号を生成して記入）** |
| J | 索引 | 参照しない（pdf\_annotate.py が出力） |

## 処理の詳細

### 1. URL収集とファイル名割り当て（パス1）

F列を2行目から順にスキャンし、セル内のURLを正規表現で抽出します。各URLは正規化（末尾スラッシュ・フラグメント除去）した上で重複判定されます。

初出のURLに対して、URLの内容に応じたプレフィックス＋通し番号のファイル名を割り当てます。

```
https://www.zabbix.com/documentation/current/en/manual/config/users
  → ZBX001

https://helpcenter.veeam.com/docs/backup/vsphere/overview.html
  → VEM001

https://example.com/docs/security-policy
  → GEN001（DEFAULTグループ）
```

既出のURLが別の行で再び現れた場合、同じファイル名が再利用されます。これにより、1つのURLに対してPDFは1つだけ生成されます。

### 2. プレフィックスの自動判別

URLに含まれるキーワードでプレフィックスが決まります。`--rules` オプションで指定します。

```
--rules "zabbix=ZBX,veeam=VEM,grafana=GRF"
```

どのルールにもマッチしないURLは `--default-prefix` で指定したDEFAULTグループに振り分けられます（デフォルト: `GEN`）。

各プレフィックスは独立した通し番号カウンタを持ちます。

```
ZBX001, ZBX002, ZBX003, ...
VEM001, VEM002, VEM003, ...
GEN001, GEN002, GEN003, ...   ← DEFAULTグループ
```

### 3. Excel書き換え（パス2）

**I列の書き込み:** 各行のF列URLに対応するファイル名を改行区切りで書き込みます。1行に複数URLがある場合は複数のファイル名が記入されます。

**G列・H列の置換:** セル内の `[URL|テキスト]` パターンを検出し、URLを対応するファイル名に置き換えます。

```
置換前: [https://www.zabbix.com/documentation/current/en/manual/config/users|By default, Zabbix has four user roles...]
置換後: [ZBX001|By default, Zabbix has four user roles...]
```

### 4. 安全な出力

元のExcelファイルは一切変更しません。処理結果は常に別ファイルに保存されます。

- `-o` 指定時: 指定したパスに保存
- `-o` 省略時: `{元ファイル名}_named_{タイムスタンプ}.xlsx` として自動生成

## 前提条件と環境構築

### 必要な環境

本スクリプトの実行には以下が必要です。

Python 3.8以上がインストールされていること。`python3 --version` で確認できます。

Excelファイルのステップ1〜2が完了していること。具体的には、人間によるA〜D列の記入と、AIエージェントによるE〜H列の記入（`prompt.md` に従った処理）が済んでいる必要があります。特にF列（URL）と、G列・H列（`[URL|テキスト]` 形式の証跡エントリ）が記入済みであることが前提です。

### Python パッケージのインストール

仮想環境の使用を推奨します。

```bash
# 仮想環境の作成と有効化（初回のみ）
python3 -m venv url2pdf-env
source url2pdf-env/bin/activate    # Linux / macOS
# url2pdf-env\Scripts\activate     # Windows

# 必要パッケージのインストール
pip install openpyxl
```

`openpyxl` はExcelファイル（.xlsx）の読み込みと書き込みに使用します。それ以外の依存パッケージはすべてPython標準ライブラリ（`argparse`、`re`、`sys`、`datetime`、`pathlib`、`collections`）のため、追加インストールは不要です。

### インストールの確認

```bash
python -c "import openpyxl; print('openpyxl:', openpyxl.__version__)"
```

バージョンが表示されれば準備完了です。

### 入力ファイルの確認

実行前に、Excelファイルが以下の状態になっていることを確認してください。

```
作業ディレクトリ/
├── assign_filenames.py      ← 本スクリプト
└── input.xlsx               ← ステップ2完了済みのExcel
                                F列: URL が記入済み
                                G列: [URL|英語テキスト] 形式で記入済み
                                H列: [URL|日本語訳] 形式で記入済み
                                I列: 空欄（本スクリプトが書き込む）
```

## 使い方

### 基本

```bash
python assign_filenames.py input.xlsx
```

出力は `input_named_20260328_150000.xlsx` のような名前で同じディレクトリに保存されます。

### 事前確認（dry-run）

```bash
python assign_filenames.py input.xlsx --dry-run
```

Excelを変更せず、URLの検出結果とファイル名の割り当て結果のみを表示します。初回実行時はまずこれで確認することを推奨します。

### ルールのカスタマイズ

```bash
python assign_filenames.py input.xlsx \
    --rules "zabbix=ZBX,veeam=VEM,grafana=GRF,microsoft=MSF" \
    --default-prefix OTH
```

### 出力先を明示

```bash
python assign_filenames.py input.xlsx -o output_final.xlsx
```

### 桁数や開始番号の変更

```bash
# 4桁で100番から開始
python assign_filenames.py input.xlsx --digits 4 --start 100
# → ZBX0100, ZBX0101, ...
```

### 列位置のカスタマイズ

Excelの列構成が異なる場合に対応できます。

```bash
python assign_filenames.py input.xlsx \
    --url-col G \
    --evidence-cols "H,I" \
    --name-col J
```

## コマンドラインオプション一覧

| オプション | デフォルト | 説明 |
|---|---|---|
| `excel` (必須) | – | 入力Excelファイル |
| `-o, --output` | 自動生成 | 出力Excelファイルのパス |
| `--rules` | `zabbix=ZBX,veeam=VEM,harvester=HRV,microsoft=MSO,kasten=VEM,kubevirt=HRV,oracle=ORA` | URL→プレフィックスの判別ルール |
| `--default-prefix` | `GEN` | どのルールにもマッチしないURLのプレフィックス |
| `--start` | `1` | 各プレフィックスの通し番号開始値 |
| `--digits` | `3` | 通し番号の桁数 |
| `--url-col` | `F` | URL列 |
| `--evidence-cols` | `G,H` | 証跡テキスト列（カンマ区切り） |
| `--name-col` | `I` | ファイル名書き込み列 |
| `--dry-run` | – | 変更せず結果のみ表示 |

## 出力例

```
============================================================
  assign_filenames – URL→ファイル名割り当て
============================================================
  入力: input.xlsx
  出力: input_named_20260328_150000.xlsx
  ルール:
    URL に 'zabbix' を含む → ZBXxxx
    URL に 'veeam' を含む → VEMxxx
    マッチなし（DEFAULT） → GENxxx
  開始番号: 1  桁数: 3
============================================================

  URL検出: 83 件（ユニーク）
    GEN (DEFAULT): 25 件
    VEM: 35 件
    ZBX: 23 件

  ┌─ GENグループ (DEFAULT) (25件)
  │  GEN001  ←  https://example.com/docs/security-policy
  │  GEN002  ←  https://example.com/docs/access-control
  │  ...
  └─

  ┌─ VEMグループ (35件)
  │  VEM001  ←  https://helpcenter.veeam.com/docs/backup/...
  │  VEM002  ←  https://helpcenter.veeam.com/docs/agents/...
  │  ...
  └─

  ┌─ ZBXグループ (23件)
  │  ZBX001  ←  https://www.zabbix.com/documentation/current...
  │  ZBX002  ←  https://www.zabbix.com/documentation/current...
  │  ...
  └─

  処理完了:
    URL（ユニーク）: 83 件
      GEN (DEFAULT): 25 件
      VEM: 35 件
      ZBX: 23 件
    I列更新: 72 行
    G/H列置換: 144 箇所
    保存先: input_named_20260328_150000.xlsx

  ※ 元ファイル 'input.xlsx' は変更されていません。
============================================================
```

## 注意事項

1行目はヘッダー行として扱われ、処理対象外です。

G列・H列の置換は `[URL|テキスト]` 形式のみが対象です。URLがブラケット外にある場合は置換されません。

URL正規化では末尾スラッシュとフラグメント（`#...`）を除去して比較します。クエリパラメータ（`?...`）は区別されます。

`--dry-run` は処理結果の事前確認用です。本番実行前に必ず確認することを推奨します。

本スクリプトはExcelの2行目からデータが存在する最終行までを自動的に処理します（`ws.max_row` を使用した動的検出）。
