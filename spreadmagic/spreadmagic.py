import shlex
import tempfile
import os
import socket
import threading
import subprocess
import time
import json
import IPython.core.magic as magic  # type: ignore  # noqa: F401
import IPython.display as display  # type: ignore  # noqa: F401


PYS_DEFAULT_VERSION = "2024.10.1"


# Google Colabで実行しているかどうかを判定
def is_google_colab():
    try:
        import google.colab  # type: ignore  # noqa: F401
        return True
    except ImportError:
        return False


# ベースディレクトリを取得
def get_basedir():
    if is_google_colab():
        if os.path.exists("/content/drive/MyDrive"):
            return "/content/drive/MyDrive/Colab Notebooks"
        else:
            return "/content"
    else:
        return os.getcwd()


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
        %%runss [width] [height] [background] [py_type] [py_conf] [js_src] [version]

    Args:
        width: iframeの幅を指定します。デフォルトは500です。
        height: iframeの高さを指定します。デフォルトは500です。
        background: iframeの背景色を指定します。デフォルトはwhiteです。
        py_type: 実行するPythonの種類。pyまたはmpyを指定します。pyはCPython互換のPyodide、mpyはMicroPytonで実行します。デフォルトはmpyです。
        py_conf: PyScriptの設定を''で囲んだJSON形式で指定します。デフォルトは{}です。
        js_src: 外部JavaScriptのURLを''で囲んだ文字列のJSON配列形式で指定します。デフォルトは[]です。
        version: PyScriptのバージョンを指定します.
    """
    iframe_execute(line, cell, False)


@magic.register_cell_magic
def genss(line, cell):
    """
    セル内のPythonコードをPyScriptを用いてiframe内で実行するために生成したHTMLを表示するマジックコマンド
    """
    iframe_execute(line, cell, True)


def iframe_execute(line, cell, viewmode):
    # 引数のパース
    args = shlex.split(line)
    width = int(args[0]) if len(args) > 0 else 500
    height = int(args[1]) if len(args) > 1 else 500
    background = args[2] if len(args) > 2 else "white"
    py_type = args[3] if len(args) > 3 else "mpy"
    py_conf = args[4] if len(args) > 4 and args[4] != "{}" else None
    js_src = args[5] if len(args) > 5 and args[5] != "[]" else None
    version = args[6] if len(args) > 6 else PYS_DEFAULT_VERSION

    options = cell.strip()

    if py_type != "py" and py_type != "mpy":
        raise ValueError("Invalid type. Use py or mpy")

    # 外部JavaSript要素を生成
    if js_src is not None:
        try:
            srcs = json.loads(js_src)
            if not isinstance(srcs, list):
                raise ValueError("Invalid JSON List format for js_src")
            js_srctag = "\n".join([f'    <script src="{src}"></script>' for src in srcs])
            js_srctag = js_srctag.rstrip("\n")
            js_srctag = f"\n{js_srctag}"
        except json.JSONDecodeError:
            raise ValueError("Invalid JSON List format for js_src")
    else:
        js_srctag = ""

    # py-config要素を生成
    if py_conf is not None:
        try:
            json.loads(py_conf)
        except json.JSONDecodeError:
            raise ValueError("Invalid JSON format for py_conf")
        py_config = f"\n    <{py_type}-config>{py_conf}</{py_type}-config>"
    else:
        py_config = ""

    # コードのHTMLテンプレート生成
    base_html = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://pyscript.net/releases/{version}/core.css" />
    <link rel="stylesheet" href="https://jsuites.net/v4/jsuites.css" type="text/css" />
    <link rel="stylesheet" href="https://bossanova.uk/jspreadsheet/v4/jexcel.css" type="text/css" />
    <script src="https://bossanova.uk/jspreadsheet/v4/jexcel.js"></script>
    <script src="https://jsuites.net/v4/jsuites.js"></script>
    <script type="module" src="https://pyscript.net/releases/{version}/core.js"></script>{js_srctag}{py_config}
    <script type="module">
        const loading = document.getElementById('loading');
        addEventListener('{py_type}:ready', () => loading.close());
        loading.showModal();
    </script>
</head>
<body style="background:{background};">
    <dialog id="loading" style="outline:none; border:none; background:transparent;">
        <div class="spinner-border" role="status"></div>
        <span class="sr-only">Loading PyScript...</span>
    </dialog>
    <div id="spreadsheet"></div>
    <script type="{py_type}">
{options}

import js
from pyscript import display, ffi

js.options = ffi.to_js(options)
call_init = js.Function("if (typeof window.options.oninit === 'function') {{ window.options.oninit(window.spreadsheet); }}")
js.spreadsheet = js.jspreadsheet(js.document.getElementById('spreadsheet'), js.window.options);
call_init()
    </script>
</body>
</html>
    """.strip()

    if viewmode:
        # HTMLを表示
        display.display(display.Pretty(base_html))

    else:
        # 一時ファイルを作成
        basedir = get_basedir()
        with tempfile.NamedTemporaryFile(mode="w", suffix=".html", delete=False, dir=basedir, encoding="utf-8") as f:
            f.write(base_html)
            temp_html_path = f.name

        # サーバーを起動
        port = find_free_port()

        # サーバーURLを取得
        url = get_server_url(port)

        # サーバーを起動
        start_server(temp_html_path, port)

        # ファイル名をURLに追加
        htmlurl = url + "/" + os.path.basename(temp_html_path)

        # IFrameを使用して表示
        display.display(display.IFrame(src=htmlurl, width=width, height=height))


# 18000番台で空いているポート番号を取得
def find_free_port(start=18000, end=18099):
    for port in range(start, end):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('localhost', port)) != 0:
                return port
    raise RuntimeError("No free ports available")


# JavaScriptを実行してプロキシURLを取得
def get_server_url(port):
    if is_google_colab():
        from google.colab.output import eval_js  # type: ignore
        url = eval_js(f"google.colab.kernel.proxyPort({port})").strip("/")
    else:
        url = f"http://localhost:{port}"

    return url


# サーバを起動する関数
def start_server_func(file_path, port):
    TIMEOUT = 30  # サーバーのタイムアウト（秒）
    server_process = subprocess.Popen(["python", "-m", "http.server", f"{port}"])
    time.sleep(TIMEOUT)
    server_process.terminate()  # プロセスの終了
    os.remove(file_path)


# サーバを別スレッドで起動
def start_server(file_path, port):
    thread = threading.Thread(target=start_server_func, args=(file_path, port))
    thread.daemon = True
    thread.start()
    time.sleep(1)  # サーバーが起動するまで待つ
