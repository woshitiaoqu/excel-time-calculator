# -*- coding: utf-8 -*-
import sys
import io
import os
import shutil
import tempfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from selenium import webdriver
from selenium.webdriver.edge.options import Options

SRC = r"C:\Users\asus\AppData\Local\Microsoft\Edge\User Data"
tmp = os.path.join(tempfile.gettempdir(), "edge_profile_copy")
if os.path.exists(tmp):
    shutil.rmtree(tmp)
os.makedirs(tmp)

print("步骤1: 复制 profile...", flush=True)
for item in os.listdir(SRC):
    s = os.path.join(SRC, item)
    t = os.path.join(tmp, item)
    if item in ("Default", "Local State", "Network"):
        try:
            if os.path.isdir(s):
                shutil.copytree(s, t)
            else:
                shutil.copy2(s, t)
        except Exception as e:
            print("  复制 %s 失败: %s" % (item, e))
print("步骤2: profile 复制完成", flush=True)

opts = Options()
for a in ["--headless=new", "--disable-gpu", "--no-sandbox",
          "--disable-dev-shm-usage", "--no-proxy-server"]:
    opts.add_argument(a)
opts.add_argument("--user-data-dir=" + tmp)
opts.add_argument("--profile-directory=Default")

print("步骤3: 启动 Edge...", flush=True)
d = webdriver.Edge(options=opts)
print("步骤4: Edge 启动成功", flush=True)
d.get("https://docs.qq.com/sheet/DWGhDUEJsek1IcGNH")
print("步骤5: 页面打开", flush=True)
import time
time.sleep(6)
print("标题:", d.title, flush=True)
ready = d.execute_script("return !!(window.SpreadsheetApp && window.SpreadsheetApp.workbook)")
print("步骤6: 就绪 =", ready, flush=True)
d.quit()
shutil.rmtree(tmp, ignore_errors=True)
