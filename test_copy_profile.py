# -*- coding: utf-8 -*-
"""测试：复制 Edge profile 到临时目录，避免锁定原 profile。"""
import sys
import io
import os
import shutil
import tempfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from selenium import webdriver
from selenium.webdriver.edge.options import Options

SRC = r"C:\Users\asus\AppData\Local\Microsoft\Edge\User Data"
# 只复制必要文件：Cookies + Local State + 相关目录
tmp = os.path.join(tempfile.gettempdir(), "edge_profile_copy")
if os.path.exists(tmp):
    shutil.rmtree(tmp)
os.makedirs(tmp)

for item in os.listdir(SRC):
    s = os.path.join(SRC, item)
    t = os.path.join(tmp, item)
    if item in ("Default", "Local State", "Network", "Preferences"):
        try:
            if os.path.isdir(s):
                shutil.copytree(s, t)
            else:
                shutil.copy2(s, t)
        except Exception as e:
            print("复制 %s 失败: %s" % (item, e))

print("临时 profile 大小:", sum(os.path.getsize(os.path.join(r, f))
      for r, _, fs in os.walk(tmp) for f in fs), "bytes")

opts = Options()
for a in ["--headless=new", "--disable-gpu", "--no-sandbox",
          "--disable-dev-shm-usage", "--no-proxy-server"]:
    opts.add_argument(a)
opts.add_argument("--user-data-dir=" + tmp)
opts.add_argument("--profile-directory=Default")

print("启动 Edge（临时 profile）...", flush=True)
d = webdriver.Edge(options=opts)
try:
    print("打开腾讯文档...", flush=True)
    d.get("https://docs.qq.com/sheet/DWGhDUEJsek1IcGNH")
    import time
    time.sleep(6)
    print("标题:", d.title)
    ready = d.execute_script(
        "return !!(window.SpreadsheetApp && window.SpreadsheetApp.workbook)")
    print("SpreadsheetApp 就绪:", ready)
finally:
    d.quit()
    shutil.rmtree(tmp, ignore_errors=True)
