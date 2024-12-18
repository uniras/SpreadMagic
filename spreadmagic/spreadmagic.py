import shlex
import IPython.core.magic as magic  # type: ignore  # noqa: F401
import pysmagic


# magic commandを登録する関数
def register_smagic():
    from IPython import get_ipython  # type: ignore  # noqa: F401
    ipy = get_ipython()
    ipy.register_magic_function(runss)
    ipy.register_magic_function(genss)
    print("Registered spread magic commands.")


# iframe内でPyScriptを実行するマジックコマンド
@magic.register_cell_magic
def runss(line, cell):
    """
    セル内のPythonコードをPyScriptを用いてiframe内で実行するマジックコマンド

    Usage:
        %%runss [width] [height] [background] [py_type] [py_val] [py_conf] [js_src] [version]

    Args:
        width: iframeの幅を指定します。デフォルトは500です。
        height: iframeの高さを指定します。デフォルトは500です。
        background: iframeの背景色を指定します。デフォルトはwhiteです。
        py_type: 実行するPythonの種類。pyまたはmpyを指定します。pyはCPython互換のPyodide、mpyはMicroPytonで実行します。デフォルトはmpyです。
        py_val: PyScriptの変数を''で囲んだJSON形式で指定します。デフォルトは'{}'です。
        py_conf: PyScriptの設定を''で囲んだJSON形式で指定します。デフォルトは'{}'です。
        js_src: 外部JavaScriptのURLを''で囲んだ文字列のJSON配列形式で指定します。デフォルトは'[]'です。
        version: PyScriptのバージョンを指定します.
    """
    args = parse_spread_args(line, cell)
    args = set_spread_args(args)
    args["htmlmode"] = False
    pysmagic.run_pyscript(args)


@magic.register_cell_magic
def genss(line, cell):
    """
    セル内のPythonコードをPyScriptを用いてiframe内で実行するために生成したHTMLを表示するマジックコマンド
    """
    args = parse_spread_args(line, cell)
    args = set_spread_args(args)
    args["htmlmode"] = True
    pysmagic.run_pyscript(args)


def run_spreadscript(args):
    args = set_spread_args(args)
    pysmagic.run_pyscript(args)


def parse_spread_args(line, cell):
    # 引数のパース
    line_args = shlex.split(line)
    args = {}
    args["width"] = line_args[0] if len(line_args) > 0 else "500"
    args["height"] = line_args[1] if len(line_args) > 1 else "500"
    args["background"] = line_args[2] if len(line_args) > 2 else "white"
    args["py_type"] = line_args[3].lower() if len(line_args) > 3 else "mpy"
    args["py_val"] = line_args[4] if len(line_args) > 4 and line_args[4] != "{}" else None
    args["py_conf"] = line_args[5] if len(line_args) > 5 and line_args[5] != "{}" else None
    args["js_src"] = line_args[6] if len(line_args) > 6 and line_args[6] != "[]" else None
    args["py_ver"] = line_args[7] if len(line_args) > 7 else "none"
    args["py_script"] = cell

    return args


def set_spread_args(args):
    # pythonコードを取得
    py_script = args.get("py_script", "")

    # 外部CSSの設定
    args["add_css"] = ["https://jsuites.net/v4/jsuites.css", "https://bossanova.uk/jspreadsheet/v4/jexcel.css"]

    # 外部JavaScriptの追加
    args["add_src"] = ["https://bossanova.uk/jspreadsheet/v4/jexcel.js", "https://jsuites.net/v4/jsuites.js"]

    # jspreadsheet実行スクリプト
    addcode = """

import js
from pyscript import display, ffi

spreadelement = js.document.createElement('div')
spreadelement.id = "spreadsheet"
js.document.body.insertBefore(spreadelement, js.document.body.firstChild)
js.options = ffi.to_js(options)
call_init = js.Function("if (typeof window.options.oninit === 'function') { window.options.oninit(window.spreadsheet); }")
js.spreadsheet = js.jspreadsheet(js.document.getElementById('spreadsheet'), js.window.options);
call_init()
"""

    # セル内のPythonコードとjspreadsheet実行スクリプトを結合
    args["py_script"] = py_script + addcode

    return args
