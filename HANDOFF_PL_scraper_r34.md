# HANDOFF_PL_scraper_r34.md

## 概要
- 作業ディレクトリ: `C:\Users\hp\Documents\Investment`
- 最新スクリプト: `pl_scraper_3files_r34.py`
- 目的: 減価償却費 (`genka`) の誤取得を是正し、損益計算書や貸借対照表由来の減価償却関連概念を採用しないようにする

## r34 の修正内容
1. EDINET preflight の認証方式を本体処理と同じに統一
- 背景:
  - `preflight_required_sources()` だけが `Ocp-Apim-Subscription-Key` ヘッダ主体で認証していた
  - 一方、通常の `edinet_fetch_docs_by_day()` / ZIP取得処理は `Subscription-Key` をクエリにも載せていた
- 変更:
  - preflight でも `Subscription-Key` をクエリとヘッダの両方に付与
  - `Subscription-Key` / `Ocp-Apim-Subscription-Key` の両ヘッダを送るように統一
- 目的:
  - preflight だけ認証方式が異なることで、実データ取得前に `401` で落ちるリスクを減らす

2. `\\w` の `SyntaxWarning` を解消
- `_extract_ticker_from_path()` の説明文字列内の `(?!\\w)` をエスケープ済み表記に修正

## r33 の修正内容
1. TDNet/EDINET のインラインXBRLから `genka` を拾う際の採用条件を厳格化
- 新規関数: `is_trusted_depr_candidate()`
- 方針:
  - `OpeCF`、`CashFlow`、`OperatingActivities` など、キャッシュ・フロー計算書またはその注記に紐づく減価償却のみ許可
  - 以下のような非CF系概念は明示的に除外
    - `DepreciationSGA`
    - `AccumulatedDepreciation`
    - `DepreciationNOE`
    - `DepreciationSegmentInformation`
    - `ReserveForAdvancedDepreciation`
    - `ReversalOfReserveForAdvancedDepreciation`

2. 定性HTMLのCF注記テーブルから取得した `genka` を最優先に変更
- `qualitative.htm` 等から抽出した `qual_metrics['genka']` を、インラインfactより常に優先
- あわせて `metric_single['genka']` を破棄し、誤った単独値が後段の補完を邪魔しないようにした

## 発生していた不具合の原因
1. 7254 ユニバンス
- 従来は `DepreciationSGA` を `genka` として採用していた
- そのため、P/L由来の `296.461` が `JB2490` に入っていた
- 本来採用すべき値は、CF注記の `2,249.752`

2. 1768 / 3690 / 4838
- 従来は `AccumulatedDepreciation` などのBS系概念を `genka` 候補として誤認していた
- これらは負値になることが多く、後段の非負チェックで捨てられる
- 結果として、CF注記に正しい数値があっても空欄のまま残っていた

## 検証結果
### フォーカス4銘柄
- 実行ファイル: `_v6_focus4_pl_v33.xlsx`
- 実行ログ: `_v6_focus4_v33_log.csv`
- 結果: `errors=0`

確認済みセル:
- 1768 行58: `JB58 = 73.529`
- 3690 行814: `IZ814 = 41.410`
- 4838 行1429: `JB1429 = 389.895`
- 7254 行2490: `JB2490 = 2249.752`

### 回帰確認8銘柄
- 対象: `1768, 3690, 4838, 7254, 3918, 6383, 7545, 175A`
- 実行ファイル: `_v6_regression_pl_v33.xlsx`
- 実行ログ: `_v6_regression_v33_log.csv`
- 結果: `errors=0`

回帰確認の意図:
- 3918: 決算期変更後のFY推定が維持されること
- 6383: gross累積補完が維持されること
- 7545: EDINET半期報告書からのgross取得が維持されること
- 175A: TDNet決算短信からのgross取得と決算期補完が維持されること

## 主要なコード位置
- バージョン更新: `pl_scraper_3files_r34.py:64`
- EDINET preflight 認証統一: `pl_scraper_3files_r34.py:4606`
- `genka` 候補の信頼性判定: `pl_scraper_3files_r34.py:3045`
- `genka` に対するCF注記優先: `pl_scraper_3files_r34.py:3300`

## 既知事項
- `python -m py_compile pl_scraper_3files_r34.py` は通過

## ファイル
- 実装: `pl_scraper_3files_r34.py`
- 引継ぎ: `HANDOFF_PL_scraper_r34.md`
- 検証出力:
  - `_v6_focus4_pl_v33.xlsx`
  - `_v6_focus4_v33_log.csv`
  - `_v6_regression_pl_v33.xlsx`
  - `_v6_regression_v33_log.csv`
