# サンプルファイル

## 概要

このディレクトリには、ワークフローの各ステップにおけるExcelの状態を示すサンプルCSVファイルが格納されています。実際の運用ではExcel（.xlsx）形式を使用しますが、GitHubでの閲覧性とdiff管理のためCSV形式で提供しています。

## ファイル一覧

| ファイル | 対応ステップ | 状態 |
|---|---|---|
| sample_step1.csv | ステップ1完了時点 | 人間がA〜D列を記入済み |
| sample_step2.csv | ステップ2完了時点 | AIエージェントがE〜H列を記入済み |
| sample_step3.csv | ステップ3完了時点 | assign_filenames.py によりI列記入・G/H列置換済み |

## 各ファイルの詳細

### sample_step1.csv — 人間による要件記入

ワークフローの出発点です。人間が記入するのはA〜D列（大分類・中分類・項番・要件）のみで、E列以降は全て空欄です。

このファイルをAIエージェントに渡し、`prompt.md` の指示に従ってE〜H列を埋めてもらうことで sample_step2.csv の状態になります。

サンプルでは、セキュリティ監査で一般的な3つの大分類（アクセス制御・データ保護・監査）にまたがる10件の要件を収録しています。

### sample_step2.csv — AIエージェントによる証跡記入

AIエージェントがE〜H列を一括で記入した状態です。この時点ではG列・H列のブラケット内にはURLがそのまま記載されており、ファイル名への置換はまだ行われていません。

```
G列の例: [https://docs.example.com/en/manual/auth/overview|Users are required to...]
H列の例: [https://docs.example.com/en/manual/auth/overview|ユーザーはシステムに...]
```

このファイルが `assign_filenames.py` の入力となります。

サンプルに含まれるURLは全てダミー（`docs.example.com`、`docs.example-backup.com`）です。実際の運用では、AIエージェントが要件に適合する実在のドキュメントURLを記入します。

### sample_step3.csv — ファイル名割り当て完了

`assign_filenames.py` の実行後の完成形です。以下の変更が適用されています。

I列（ファイル名）に通し番号付きファイル名が書き込まれています。G列・H列のブラケット内のURLがファイル名に置換されています。

```
G列の例: [EVD001|Users are required to...]   ← URLがEVD001に置換
H列の例: [EVD001|ユーザーはシステムに...]     ← 同上
I列の例: EVD001                                ← ファイル名が記入
```

このファイルが `url2pdf.py`（ステップ4）および `pdf_annotate.py`（ステップ5）の入力となります。

## サンプルの設計意図

### 同一URLの複数行からの参照

実際の監査では、1つのドキュメントページが複数の要件の証跡として共通的に参照されることが頻繁にあります。サンプルではこのパターンを3箇所に含めています。

| 項番 | URL | 割当ファイル名 | 意図 |
|---|---|---|---|
| 1.2.1 と 1.2.2 | `docs.example.com/en/manual/auth/roles` | 共に EVD003 | 同一URLから異なる証跡テキストを抽出するケース |
| 2.1.1 と 2.1.2 | `docs.example.com/en/manual/security/encryption` | 共に EVD004 | 同上 |
| 3.1.1 と 3.1.2 | `docs.example.com/en/manual/audit/logging` | 共に EVD005 | 同上 |

このパターンにより、以下の動作を確認できます。

`assign_filenames.py` では、同一URLに対して同一のファイル名が割り当てられ、重複したPDF生成が防止されます。`url2pdf.py` では、I列の重複排除により同じURLからPDFが2回生成されることはありません。`pdf_annotate.py` では、同一PDF内の異なるテキスト箇所にそれぞれハイライトが付与され、項番ラベルは同一ページ内で重複排除されます。

### 異なるドメインによるプレフィックス振り分け

サンプルには2つの異なるドメインのURLを含めています。

| ドメイン | 想定プレフィックス | 該当行 |
|---|---|---|
| `docs.example.com` | EVD（デフォルト） | 項番 1.1.1〜1.2.2、2.1.1〜2.1.2、3.1.1〜3.1.2 |
| `docs.example-backup.com` | BKP | 項番 2.2.1〜2.2.2 |

`assign_filenames.py` を以下のように実行することで、URLのドメインに応じたプレフィックスの振り分けが確認できます。

```bash
python assign_filenames.py sample_step2.csv \
    --rules "example-backup=BKP" \
    --default-prefix EVD \
    --dry-run
```

### E列（回答）の簡潔さ

E列の回答は意図的に短く簡潔にしています。これは `prompt.md` で指示している「もっともシンプルな内容」という方針に沿ったものです。回答はG列の証跡テキストの要約であり、詳細な説明はG列・H列の証跡テキスト自体が担います。

### G列・H列のテキスト長

G列の証跡テキストは1〜2文程度の長さに収めています。これは `pdf_annotate.py` のパターンマッチ検索で確実にヒットさせるための実用的な長さです。短すぎると誤マッチのリスクがあり、長すぎるとPDF化時の微妙な空白・改行の差異によりマッチに失敗する可能性が高くなります。

## サンプルを使った動作確認

サンプルファイルを使ってワークフローの一部を試すことができます。ただし、URLはダミーのため `url2pdf.py`（ステップ4）と `pdf_annotate.py`（ステップ5）は実際には動作しません。

```bash
# ステップ3の動作確認（assign_filenames.py）
# sample_step2.csv をExcelに変換してから実行するか、
# スクリプトがCSVに対応していない場合は .xlsx に変換してください

# dry-runで結果を確認
python assign_filenames.py sample_step2.xlsx \
    --rules "example-backup=BKP" \
    --default-prefix EVD \
    --dry-run
```

出力結果が sample_step3.csv の状態と一致すれば、`assign_filenames.py` が正しく動作しています。

## 列構成リファレンス

全ステップ共通の列構成です。各ステップで記入される列が異なります。

| 列 | ヘッダー名 | step1 | step2 | step3 | 記入者 |
|---|---|---|---|---|---|
| A | 大分類 | ✓ | ✓ | ✓ | 人間 |
| B | 中分類 | ✓ | ✓ | ✓ | 人間 |
| C | 項番 | ✓ | ✓ | ✓ | 人間 |
| D | 要件 | ✓ | ✓ | ✓ | 人間 |
| E | 回答 | — | ✓ | ✓ | AI |
| F | URL | — | ✓ | ✓ | AI |
| G | 証跡テキスト原文 | — | ✓（URL形式） | ✓（ファイル名形式） | AI → スクリプト置換 |
| H | 日本語訳 | — | ✓（URL形式） | ✓（ファイル名形式） | AI → スクリプト置換 |
| I | ファイル名 | — | — | ✓ | assign_filenames.py |
| J | 索引 | — | — | — | pdf_annotate.py |
