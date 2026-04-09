# PL Scraper (`pl_scraper_3files.py`)

日本上場企業の **PL（損益計算書）指標** を EDINET / TDNet / Kabutan から自動取得し、Excel テンプレートを一括更新するスクレイパー。

---

## 概要

| 項目 | 内容 |
|------|------|
| 対応ファイル数 | 最大 3 ファイル同時（単独四半期 / 半期累積 / 年次） |
| データソース | EDINET（最優先）→ TDNet LIVE → TDNet GitHub アーカイブ → Kabutan |
| 値の単位 | 百万円（千円単位の開示は自動変換） |
| 上書きポリシー | **空欄のみ**書き込み。既存値は保護（±2 百万円以上の乖離で橙 WARN） |
| 対応銘柄コード | 4 桁数字（例: 9760）および英数混合（例: 142A）の新形式 |
| バージョン | r36 |

---

## 取得指標

| 内部キー | 日本語名 | 対象期間（左ブロック） |
|----------|----------|----------------------|
| `keijo` | 経常利益 | FY2018Q1〜FY2025Q4（列 S:AX） |
| `uriage` | 売上高合計 | FY2018Q1〜FY2025Q4（列 AY:CD） |
| `saishu` | 親会社株主に帰属する当期純利益 | FY2018Q1〜FY2025Q4（列 CE:DJ） |
| `genka` | 減価償却費 | FY2023Q1〜FY2025Q4（列 DK:DV） |
| `gross` | 売上総利益 | FY2022Q1〜FY2025Q4（列 DW:EL） |
| `sga` | 販売費及び一般管理費 | FY2023Q1〜FY2025Q4（列 EM:EX） |

右ブロック（列 EZ〜KE）には**累積値**が格納されます（左ブロックの単独四半期値とは独立）。

---

## セル着色ルール

| 色 | 意味 |
|----|------|
| **黄色**（`#FFF2CC`） | 今回スクリプトが新規入力した値 |
| **橙色**（`#F4B183`） | 既存値と取得値の差が `warn_tolerance`（デフォルト ±2 百万円）超 → WARN |

- `EY` 列（WARN フラグ列）: 直近 5 四半期（FY2024Q4〜FY2025Q4）に 1 件でも橙 WARN があれば `EY=1`。

---

## ワークブック構成（想定）

```
データ取得_PL.xlsx        ← 単独四半期（--input）
データ取得_半期累積.xlsx  ← 半期累積（--half-input）
年次_データ取得.xlsx      ← 年次（--annual-input）
```

- データは 4 行目から開始（`--start-row 4`）。
- 各行の **A 列または B 列**に証券コード（ticker）が格納されていることを前提とします。

---

## インストール

```bash
pip install requests beautifulsoup4 openpyxl lxml
# WAF 回避が必要な場合（任意）
pip install curl_cffi
```

Python 3.9 以上を推奨。

---

## 使い方

### 基本（3 ファイル同時実行）

```bash
python pl_scraper_3files.py \
  --input       データ取得_PL.xlsx          --output      _out_pl.xlsx \
  --half-input  データ取得_半期累積.xlsx    --half-output _out_half.xlsx \
  --annual-input 年次_データ取得.xlsx       --annual-output _out_annual.xlsx \
  --edinet-api-key <YOUR_EDINET_API_KEY> \
  --log-csv _out_log.csv
```

Windows（コマンドプロンプト）では `\` を `^` に置き換えてください。

### 単独ファイルのみ

```bash
python pl_scraper_3files.py \
  --input データ取得_PL.xlsx --output _out_pl.xlsx \
  --edinet-api-key <YOUR_EDINET_API_KEY>
```

### 特定 ticker のみ処理（テスト用）

```bash
python pl_scraper_3files.py \
  --input データ取得_PL.xlsx --output _out_pl.xlsx \
  --edinet-api-key <YOUR_EDINET_API_KEY> \
  --tickers "4395,4912,436A,5830"
```

### 処理行数を制限（スモークテスト）

```bash
python pl_scraper_3files.py \
  --input データ取得_PL.xlsx --output _out_pl.xlsx \
  --edinet-api-key <YOUR_EDINET_API_KEY> \
  --limit 5
```

---

## 全 CLI オプション

| オプション | デフォルト | 説明 |
|-----------|-----------|------|
| `--input` | **必須** | 単独四半期 xlsx（入力） |
| `--output` | **必須** | 単独四半期 xlsx（出力） |
| `--half-input` | `` | 半期累積 xlsx（入力） |
| `--half-output` | 自動 | 半期累積 xlsx（出力） |
| `--annual-input` | `` | 年次 xlsx（入力） |
| `--annual-output` | 自動 | 年次 xlsx（出力） |
| `--edinet-api-key` | 環境変数 `EDINET_API_KEY` / `.env` | EDINET API キー |
| `--tickers` | `` | カンマ区切り ticker ホワイトリスト |
| `--tickers-file` | `` | ticker リストのテキスト/CSV ファイル |
| `--limit` | `0`（全行） | 処理行数上限（テスト用） |
| `--start-row` | `4` | データ開始行番号 |
| `--warn-tolerance` | `2.0` | WARN 判定しきい値（百万円） |
| `--log-csv` | `` | セル単位ログ CSV の出力先 |
| `--progress-every` | `10` | 進捗表示の間隔（社数） |
| `--max-filings-per-ticker` | `6` | 1 ticker あたり最大ダウンロード件数 |
| `--error-dump` | `single` | エラーダンプ方式（`single` / `per_row` / `none`） |
| `--error-dump-file` | `` | エラーダンプ出力先（省略時は自動命名） |
| `--legacy-v5-40-single` | off | 旧単独値ポリシーを使用（非推奨） |
| `--tdnet-github-disable` | off | TDNet GitHub アーカイブを無効化 |
| `--tdnet-github-owner` | `yukizi1113` | TDNet アーカイブの GitHub オーナー |
| `--tdnet-github-repo` | `tdnet` | TDNet アーカイブのリポジトリ名 |
| `--tdnet-github-branch` | `main` | ブランチ名 |
| `--tdnet-github-xbrl-dir` | `XBRL` | XBRL ZIP が格納されているディレクトリ |
| `--tdnet-github-from-date` | `2025-12-17` | アーカイブ参照の開始日（YYYY-MM-DD） |

---

## データソースの詳細

### 1. EDINET（最優先）
- 金融庁の電子開示システム。XBRL ZIP をダウンロードして解析。
- 直近 365 日分を対象。
- API キーは `--edinet-api-key` または環境変数 `EDINET_API_KEY` で指定。

### 2. TDNet LIVE
- 東証の適時開示情報閲覧サービス（`release.tdnet.info`）。
- 直近 180 日分を対象（公開期間が短いため）。
- XBRL ZIP を直接ダウンロードして解析。

### 3. TDNet GitHub アーカイブ
- TDNet LIVE の保存期間切れ分を補完するユーザー管理 GitHub リポジトリ。
- GitHub REST API（git trees）で ZIP を列挙し `raw.githubusercontent.com` から取得。
- デフォルトで有効（`--tdnet-github-disable` で無効化）。

### 4. Kabutan（補完）
- スクレイピングによる補完ソース。EDINET/TDNet で取得できなかったデータに使用。
- `r36` 以降、Kabutan 側の一時的な `504` / タイムアウトでは全体停止せず、その ticker の Kabutan 取得だけをスキップして継続。

---

## 出力ファイル

| ファイル | 内容 |
|---------|------|
| `--output` で指定した xlsx | 更新済み単独四半期ワークブック |
| `--half-output` | 更新済み半期累積ワークブック |
| `--annual-output` | 更新済み年次ワークブック |
| `--log-csv` | セル単位のアクション・サマリーログ |
| `*_errors.txt` | エラー詳細（`--error-dump=single` 時） |
| `__pl_meta` シート | 各セルのソース・優先度等のプロベナンス（ワークブック内非表示シート） |

`r36` 以降、fatal abort が起きた場合でも、その時点までの PL ワークブックと partial log を best-effort で保存します。

---

## 環境変数

`.env` ファイルに以下を設定することで CLI オプションを省略できます。`r35` 以降は、`--edinet-api-key` / 環境変数 / ワークスペース直下 `.env` の順で EDINET API キー候補を解決します。

```dotenv
EDINET_API_KEY=your_edinet_api_key_here
GH_TOKEN=your_github_token_here          # TDNet GitHub アーカイブの認証に使用
MAX_ABS_MILLION_SANITY=1000000000        # 異常値ガード（百万円単位）
```

---

## WARN（橙）とデータ品質

- スクリプトは**既存値を絶対に上書きしません**。
- 取得値と既存値の差が `warn_tolerance` を超えた場合にのみ橙色でマークします。
- `__pl_meta` 隠しシートにより、前回実行時のソース・優先度が記録されるため、TDNet の保存期間切れ後も**誤 WARN を防止**できます。

---

## バージョン履歴（抜粋）

| バージョン | 主な変更 |
|-----------|---------|
| r36 | Kabutan を optional fallback として扱い、個別 `504` で全体停止しないよう修正。fatal abort 時の partial save を追加 |
| r35 | EDINET API キーを `--edinet-api-key` / 環境変数 / `.env` から自動解決し、preflight で通ったキーを本処理へ引き継ぐよう修正 |
| r34 | EDINET preflight の認証方式を本処理と統一し、`Subscription-Key` を query + header の両方で送るよう修正 |
| r33 | 減価償却費 (`genka`) は CF / CF注記由来のみ採用。`DepreciationSGA` や `AccumulatedDepreciation` の誤採用を遮断し、定性HTMLのCF注記値を最優先化 |
| r32 | EDINET 半期報告書・有価証券報告書の四半期判定フォールバックを追加し、半期gross取得漏れを修正 |
| r31 | gross の累積値を、既存シート上の単独値から条件付きで補完する処理を追加 |
| r30 | 決算期変更銘柄での FY 月推定を、`C:R` の右端有効セルベースに変更 |
| r25 | TDNet GitHub アーカイブ対応（保存期間切れ補完） |
| r24 | `ExtraordinaryProfit` 誤採用修正、銀行 BNK 系除外、`ProfitLoss` の `saishu` マッピング追加、CAPEX 合算対応 |
| r23 | `saishu` で `SummaryOfBusinessResults` 誤採用を修正（false WARN 解消） |
| r20 | iXBRL コンテキスト適合を最優先に変更（連結/非連結の誤採用を修正） |

詳細は [`HANDOFF_PL_scraper.md`](./HANDOFF_PL_scraper.md) を参照。過去版は git history に残ります。

---

## ライセンス

個人利用・社内利用目的。再配布・商用利用は別途確認してください。
