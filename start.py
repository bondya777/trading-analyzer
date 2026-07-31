#!/usr/bin/env python3
import http.server
import socketserver
import webbrowser
import os
import threading
import time

PORT = 7100
DIRECTORY = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist")

class MyHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

def open_browser():
    time.sleep(1)
    url = f"http://localhost:{PORT}/"
    print(f"正在打开浏览器: {url}")
    webbrowser.open(url)

if __name__ == "__main__":
    if not os.path.exists(DIRECTORY):
        print(f"错误: 找不到 dist 目录: {DIRECTORY}")
        print("请确保 dist 文件夹存在")
        input("按回车退出...")
        exit(1)

    print(f"=" * 50)
    print(f"交易数据分析平台 - 本地服务器")
    print(f"=" * 50)
    print(f"服务器地址: http://localhost:{PORT}/")
    print(f"文件目录: {DIRECTORY}")
    print(f"=" * 50)
    print(f"按 Ctrl+C 停止服务器")
    print(f"=" * 50)

    # 自动打开浏览器
    threading.Thread(target=open_browser, daemon=True).start()

    with socketserver.TCPServer(("", PORT), MyHTTPRequestHandler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n服务器已停止")
