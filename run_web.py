# -*- coding: utf-8 -*-
"""
راه‌اندازی برنامه: یک سرور محلی روشن می‌کند و صفحه را در مرورگر باز می‌کند.
برای بستن برنامه، این پنجره را ببندید.
"""
import socket
import threading
import webbrowser

from web_app import app


def free_port():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


if __name__ == "__main__":
    port = free_port()
    url = f"http://127.0.0.1:{port}"
    print("برنامه در حال اجراست:", url)
    print("برای بستن برنامه، این پنجره را ببندید.")
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=port, debug=False)
