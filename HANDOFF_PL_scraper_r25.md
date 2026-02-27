# HANDOFF_PL_scraper_r24.md

このファイルは「PLスクレイパー（3ファイル対応）」の文脈・仕様・修正履歴を引き継ぐためのハンドオフです。

## 目的（不変）
- EDINET優先・TDNet補完で PL 指標を取得し、Excel テンプレ（単独 / 半期累積 / 年次の3ファイル）を更新する。
- **入力が正しいセルには WARN（橙）が付かず、誤りのみ WARN になる**仕組みを堅牢化する。
- ticker 個別分岐は原則禁止（例外を増やさない）。

## 入出力（現行）
- input:
  - `データ取得_PL.xlsx`（単独）
  - `データ取得_半期累積.xlsx`（半期累積）
  - `年次_データ取得.xlsx`（年次）
- output:
  - `--output`（単独の出力）
  - `--half-output`（半期累積の出力。未指定なら入力名ベースで自動）
  - `--annual-output`（年次の出力。未指定なら入力名ベースで自動）
  - `--log-csv`（セル単位ログ）
  - `*_errors.txt`（エラー詳細。--error-dump=single の場合）

## 重要仕様（要点）
### 1) 既存セルの扱い
- **空欄セルのみ**値を入力（黄色）。
- 既存値があるセルは上書きしない。
- 既存値と取得値の差が **abs_diff > warn_tolerance（デフォルト±2, 百万円単位）** の場合:
  - **上書きせず**既存値を保持
  - セルを **橙色 WARN** にする（ログにも残す）

### 2) 優先順位（不変）
- 一次情報（EDINET / TDNet）を優先し、Kabutan は補完用途。

---

## r20 の重要修正（HALFで ProfitLossBeforeTax があるのに NonConsolidatedMember の OrdinaryIncome を誤採用）
- 原因: fact採用で name 優先が勝ち、コンテキスト適合（連結/非連結・Member/非Member）が後回し。
- 対応: 採用スコアの並びを変更し、**コンテキスト適合を最優先**にした。
  - 連結優先（prefer_consolidated=True のとき NonConsolidated を不利に）
  - Member/Segment を不利に
  - その上で tag-name priority を使う
- 重要: **ConsolidatedMember は Member ペナルティ対象外**（通常軸のため）。

## r23 の重要修正（HALFの最終利益で SummaryOfBusinessResults を誤採用して false WARN）
- 原因: `saishu` に Summary ペナルティが無く、KPI表の fact を拾ってしまう。
- 対応: `saishu` に `summaryofbusinessresults` のペナルティを導入し、本表を優先。

---

## r24 の追加修正（2026-02-20）
今回の主目的は **正しい値に対する false WARN の除去**（ただし誤り検知は維持）です。

### 1) 「ExtraordinaryProfit」を「ordinaryprofit」と誤判定する部分文字列トラップの解消（4395など）
- 症状: `ExtraordinaryProfit...` が local-name 内に `ordinaryprofit` を含むため、`keijo` と誤マップ → 誤採用 → 入力が正しいのに WARN。
- 対応: `metric_from_qname()` で **extraordinary + ordinary*** の場合は `keijo` にマップしない（ガードを先に置く）。

### 2) 銀行の `OrdinaryIncomeBNK`（経常収益）を `keijo` から除外（5830/5831など）
- 症状: 銀行で `OrdinaryIncomeBNK`（収益系）と `OrdinaryIncomeLoss/OrdinaryProfitLoss`（利益系）が併存し、同点タイブレークで収益側を拾って false WARN。
- 対応: `metric_from_qname()` で **BNK 系の OrdinaryIncome を `keijo` から除外**（ただし `..Loss/..ProfitLoss` は利益概念なので `keijo` に残す）。

### 3) `ProfitLoss` / `NetIncomeLoss` を `saishu` として扱う（436Aなど）
- 症状: 一部提出者（非連結/小規模/一部IFRS）で最終利益が `ProfitLoss` で出ており、従来 `saishu` にマップされず誤採用・WARN。
- 対応: `metric_from_qname()` に **保守的な exact-match で `ProfitLoss`/`NetIncomeLoss` 等を `saishu` に追加**。
- 追加: `parse_zip_metrics()` のスコアリングで **NonConsolidated の場合は AttributableToOwnersOfParent を軽く不利**にして、`ProfitLoss` を優先できるようにした。

### 4) 年次CFの CAPEX（有形固定資産取得）を「PPE + 賃貸固定資産/その他固定資産」合算で合うように（2212/5982/6737）
- 症状: CF の「有形固定資産の取得による支出」が、提出者によって
  - PPE
  - 賃貸固定資産（real estates for rent / rental fixed assets）
  - その他の固定資産（other noncurrent/fixed assets / increase in other assets）
 などに分割される。Excel が合算値（正）でも、単体PPEだけ拾って WARN。
- 対応: `parse_zip_metrics()` で **同一 contextRef 内の PPE + (rent/other) を検出して `SYNTHCapexPpeSum` を生成**し、`capex_ppe` はこの合算を最優先で採用。

### 5) ConsolidatedMember 判定の頑健化（アンダースコア無し contextRef 対応）
- 症状: contextRef の表記揺れ（`...ConsolidatedMember...` が `_` で区切られない等）で、ConsolidatedMember を誤って Member ペナルティに入れてしまい誤採用の原因になり得る。
- 対応: Member ペナルティは **substring ガード**で判定し、
  - `NonConsolidatedMember` と Segment 系のみを明確にペナルティ
  - `ConsolidatedMember` はペナルティ対象外
 となるようにした。

---

## 実行コマンド例（3ファイル同時）
```bash
python pl_scraper_3files_r24.py ^
  --input データ取得_PL.xlsx --output _out_pl.xlsx ^
  --half-input データ取得_半期累積.xlsx --half-output _out_half.xlsx ^
  --annual-input 年次_データ取得.xlsx --annual-output _out_annual.xlsx ^
  --edinet-api-key <YOUR_KEY> ^
  --log-csv _out_log.csv
```

## 部分実行（ticker指定）
- `--tickers "4395,4912,436A,5830,5831,2212,5982,6737"`

## 作成物（このチャットの最新版）
- `pl_scraper_3files_r24.py`
- `HANDOFF_PL_scraper_r24.md`


## r25 変更点（TDNet GitHub アーカイブ対応）

### 背景
- TDNet（release.tdnet.info）は公開画面の仕様上、過去の開示を一定期間（例：直近約35日など）しか遡れないことがあります。
- ユーザー管理の GitHub リポジトリ（`yukizi1113/tdnet`）に、2025-12-17 以降の TDNet XBRL ZIP / PDF が保管されているため、
  TDNet LIVE で取得できない期間の補完として利用できるようにしました。

### 実装
- 追加ソース：**TDNet GitHub archive**（扱いは「TDNet」と同一優先度）
- GitHub REST API（git trees）で `XBRL/` 配下の `.zip` を列挙し、`raw.githubusercontent.com` から ZIP を直接ダウンロードして既存の `parse_zip_metrics()` に渡します。
- PDF（`tekigikaizi/`）は本ツールでは解析しません（XBRL ZIP のみ）。
- LIVE（TDNet）と GitHub の両方が落ちている場合のみ、TDNet ソース不達として fail-fast します。

### CLI
- デフォルトで GitHub アーカイブは有効です。
- 無効化：`--tdnet-github-disable`
- 参照先変更：
  - `--tdnet-github-owner`（default: yukizi1113）
  - `--tdnet-github-repo`（default: tdnet）
  - `--tdnet-github-branch`（default: main）
  - `--tdnet-github-xbrl-dir`（default: XBRL）
  - `--tdnet-github-from-date`（default: 2025-12-17）

### 注意
- GitHub 側のファイル名/パスから ticker と日付を推定してインデックス化します。
  もしファイル名に ticker が含まれない構造の場合、インデックスに載らないため、ZIP 命名規則の調整（ticker を含める等）が必要です。
