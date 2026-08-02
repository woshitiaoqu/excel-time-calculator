# -*- coding: utf-8 -*-
"""探索 WPS ActiveWorkbook 的属性和方法。"""
import sys
import io
import time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import tencent

prof = tencent._pick_profile_for_url("https://www.kdocs.cn/l/slavMEqSn2ig")
d = tencent._create_driver(profile=prof)
try:
    d.set_page_load_timeout(60)
    d.get("https://www.kdocs.cn/l/slavMEqSn2ig")
    time.sleep(12)
    print("标题:", d.title)
    info = d.execute_script("""
        try {
          var app = window.WPSOpenApi.Application;
          var wb = app.ActiveWorkbook;
          var keys = Object.getOwnPropertyNames(wb);
          return JSON.stringify(keys.slice(0, 60));
        } catch(e) { return 'ERR:' + e.message; }
    """)
    print("ActiveWorkbook 属性:", info)
finally:
    d.quit()
