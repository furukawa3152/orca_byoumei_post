# ORCA 病名一括登録スクリプト

`Dr.入力病名一覧.xlsx` に入力された病名データを読み込み、ORCA の病名登録 API に1件ずつ送信するためのスクリプトです。

## 概要

このリポジトリでは、`deseasename_post.py` を実行することで、同じフォルダに置かれた Excel ファイルから病名情報を読み取り、ORCA API に XML 形式で POST します。

現在の実装では、Excel の1行を1件の病名登録として扱います。複数行ある場合は上から順に送信し、ORCA 側への負荷を避けるため、各送信後に `0.3` 秒待機します。

## ファイル構成

- `deseasename_post.py`
  - Excel 読み込み、データ整形、XML 作成、ORCA API への POST を行うメインスクリプトです。

- `Dr.入力病名一覧.xlsx`
  - 登録対象の病名一覧を入力する Excel ファイルです。
  - `deseasename_post.py` と同じフォルダに配置します。

## 必要な Python ライブラリ

以下のライブラリを使用します。

- `openpyxl`
  - Excel ファイルの読み込みに使用します。

- `requests`
  - ORCA API への HTTP POST に使用します。

- `xmltodict`
  - ORCA API から返ってきた XML レスポンスを辞書に変換するために使用します。

インストール例:

```bash
pip install openpyxl requests xmltodict
```

## Excel ファイル仕様

読み込み対象ファイルは、スクリプトと同じフォルダにある `Dr.入力病名一覧.xlsx` です。

現在の実装では、先頭シートの1行目をヘッダーとして読み込みます。2行目以降が登録対象データです。

必要な列:

- `ptid`
  - ORCA の患者IDです。
  - XML の `Patient_ID` に入ります。
  - 空の場合、その行はスキップされます。

- `byomei`
  - 登録する病名です。
  - XML の `Disease_Name` に入ります。
  - 空の場合、その行はスキップされます。

- `sryymd`
  - 病名の開始日です。
  - XML の `Disease_StartDate` に入ります。
  - Excel の日付型、または `YYYY-MM-DD` 形式の文字列を想定しています。
  - 空の場合、その行はスキップされます。

- `tenkimei`
  - 転帰名です。
  - `Disease_OutCome` に入れる値の判定に使用します。

- `tenkikubun`
  - 転帰区分です。
  - `tenkimei` が既知の変換対象ではない場合の予備値として使用します。

- `tenkiymd`
  - 転帰日です。
  - XML の `Disease_EndDate` に入ります。
  - 空の場合は空文字のまま送信します。

任意で存在している列:

- `tbl_ptinf::name`
  - 患者氏名などの確認用列として存在します。
  - 現在のスクリプトでは ORCA API 送信には使用していません。

## 転帰の変換仕様

`tenkimei` の値に応じて、ORCA に送る `Disease_OutCome` を変換します。

- `tenkimei` が `治癒` の場合
  - `Disease_OutCome` には `3` を送信します。

- `tenkimei` が `中止（転医）` の場合
  - `Disease_OutCome` には `N` を送信します。

- 上記以外の場合
  - `tenkikubun` の値をそのまま `Disease_OutCome` に送信します。

- `tenkimei` と `tenkikubun` がどちらも空の場合
  - `Disease_OutCome` は空文字で送信します。

## ORCA API 送信先

現在の送信先は `deseasename_post.py` 内で以下のように定義されています。

```python
ORCA_DISEASE_URL = "http://ormaster:ormaster@172.16.123.100:8000/api/orca22/diseasev3"
```

送信先や認証情報を変更する場合は、この値を変更してください。

## 送信される主な XML 項目

Excel の値は、主に以下の XML 項目に反映されます。

- `Patient_ID`
  - Excel の `ptid`

- `Base_Month`
  - `sryymd` の先頭7文字から作成した年月
  - 例: `2026-05-08` の場合は `2026-05`

- `Perform_Date`
  - 明示指定がなければ `sryymd` と同じ日付

- `Department_Code`
  - 現在は固定で `03`

- `Disease_Name`
  - Excel の `byomei`

- `Disease_StartDate`
  - Excel の `sryymd`

- `Disease_EndDate`
  - Excel の `tenkiymd`

- `Disease_OutCome`
  - Excel の `tenkimei` または `tenkikubun` から作成

- `Disease_Class`
  - 現在は固定で `Auto`

## 実行方法

1. `deseasename_post.py` と同じフォルダに `Dr.入力病名一覧.xlsx` を置きます。

2. 必要なライブラリをインストールします。

```bash
pip install openpyxl requests xmltodict
```

3. スクリプトを実行します。

```bash
python deseasename_post.py
```

4. 送信に成功した行は、以下のような形式で標準出力に表示されます。

```text
posted row 2: 123514 急性咽頭炎
posted row 3: 123514 アレルギー性鼻炎
```

## 実行間隔

ORCA API への連続送信を避けるため、各行の送信後に `0.3` 秒待機します。

待機時間は `post_diseases_from_excel()` の `interval_seconds` で変更できます。

```python
post_diseases_from_excel(interval_seconds=0.3)
```

## スキップされる行

以下のいずれかが空の場合、その行は送信されずにスキップされます。

- `ptid`
- `byomei`
- `sryymd`

スキップ時は、以下のようなメッセージが表示されます。

```text
skip row 5: ptid/byomei/sryymd のいずれかが空です
```

## 動作確認

構文チェック:

```bash
python -m py_compile deseasename_post.py
```

Excel 読み取り結果だけを確認したい場合:

```bash
python - <<'PY'
from deseasename_post import read_diseases_from_excel

for item in read_diseases_from_excel():
    print(item)
PY
```

この確認では ORCA API への POST は行われません。

## 注意点

- 実行すると ORCA API に実際に病名登録リクエストを送信します。
- 本番環境で実行する前に、Excel の患者ID、病名、開始日、転帰、転帰日を必ず確認してください。
- `Dr.入力病名一覧.xlsx` のファイル名を変更した場合は、`deseasename_post.py` の `EXCEL_PATH` も変更してください。
- `requests` と `xmltodict` は `byoumei_post()` 実行時に読み込まれます。そのため、Excel の読み取り確認だけであれば未インストールでも可能ですが、実際のPOST時には必要です。
- 病名や患者IDなど、XML に入る値は `html.escape()` によりエスケープして送信します。
- ORCA API のレスポンスは XML から辞書へ変換され、JSON 文字列として関数から返されます。

## 現在の既定値

- Excel ファイル名: `Dr.入力病名一覧.xlsx`
- ORCA API: `http://ormaster:ormaster@172.16.123.100:8000/api/orca22/diseasev3`
- 診療科コード: `03`
- 病名区分: `Auto`
- POST 間隔: `0.3` 秒

