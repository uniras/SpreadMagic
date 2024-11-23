# SpreadSheet Magic Command

## 概要

Jypyter(notebook/lab)・VSCodeまたはGoogle Colabでjspresdsheetを使ったコードセルのPythonコードをPyScriptを使ってiframe(ブラウザ)上で実行するマジックコマンドです。

## 使い方

### マジックコマンドの追加

コードセルに以下のコードを貼り付けて実行しマジックコマンドを登録してください。カーネルやランタイムを再起動する度に再実行する必要があります。

```python
%pip install -q -U pysmagic spreadmagic
from spreadmagic import register_smagic

register_smagic()
```

### マジックコマンドの使い方

コードセルの冒頭に以下のようにマジックコマンドを記述してください。実行するとアウトプットにiframeが表示されてその中でコードセルのコードがPyScriptで実行されます。

```python
%%runss 800 400 white

import js

# イベントハンドラのサンプル
def onselection(instance, x1, y1, x2, y2):
    display(f"Selected: {x1}, {y1}, {x2}, {y2}")

# カスタムセル関数のサンプル
def DOUBLE(x):
    return x * 2

# カスタムセル関数の登録
js.DOUBLE = DOUBLE

# jspreadsheet.jsに渡すオプションの設定
options = {
    # データ
    "data": [
        ["2022/10/24", "文房具", 480],
        ["2022/10/26", "外食費", 1390],
        ["2022/10/27", "外食費(2倍)", "=DOUBLE(C2)"],
    ],

    # カラム書式の定義
    "columns": [
        { "type": "calendar", "title": "日付", "width": 120, "options": { "format": "YYYY/MM/DD" } },
        { "type": "text", "title": "項目", "width": 300 },
        { "type": "numeric", "title": "出金", "width": 200, "mask":"#,##" },
    ],

    # イベントハンドラの登録
    "onselection": onselection,
}
```

### グローバル変数

PyScriptから以下の変数にアクセスできます。

- 別のセルで設定したグローバル変数(_で始まる変数名やJSONに変換できないものは除く)
- マジックコマンドの引数py_valで設定した変数
- width: iframeの幅(マジックコマンドの引数で指定した幅)
- height: iframeの高さ(マジックコマンドの引数で指定した高さ)

この変数はjs.pysオブジェクトを介してアクセスできます。
変数名が衝突した場合は上記リストの順に上書きされて適用されます。

### マジックコマンド

#### %%runss

セル内のjspreadsheet.jsを使ったPythonコードをPyScriptを用いてiframe内で実行するマジックコマンド

```juypyter
%%runss [width] [height] [background] [py_type] [py_val] [py_conf] [js_src] [py_ver]
```

- width: iframeの幅を指定します。デフォルトは500です。
- height: iframeの高さを指定します。デフォルトは500です。
- background: iframeの背景色を指定します。デフォルトはwhiteです。
- py_type: 実行するPythonの種類。pyまたはmpyを指定します。mpyはMicroPyton、pyはCPython互換のPyodideで実行します。デフォルトはmpyです。グローバルモードのときはmpy固定です。
- py_val: PyScriptの変数を''で囲んだJSON形式で指定します。デフォルトは'{}'です。
- py_conf: PyScriptの設定を''で囲んだJSON形式で指定します。デフォルトは{}です。
- js_src: 外部JavaScriptのURLを''で囲んだ文字列のJSON配列形式で指定します。デフォルトは[]です。
- py_ver: PyScriptのバージョンを指定します.

#### %%genss

セル内のjspreadsheet.jsを使ったPythonコードをPythonコードからブラウザで実行可能な単一HTMLを生成するマジックコマンド。オプションはrunssと同じです。

```juypyter
%%genss [width] [height] [background] [py_type] [py_conf] [js_src] [py_ver]
```

## 参考

[とほほのJspreadsheet入門](https://www.tohoho-web.com/ex/jspreadsheet.html)
