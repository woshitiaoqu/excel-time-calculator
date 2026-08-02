# -*- coding: utf-8 -*-
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import tencent

print("检测 profile...")
prof = tencent._find_logged_in_profile(["kdocs.cn", "wps.cn", "kingsoft"])
print("profile:", prof)

print("尝试用 profile 启动 driver...")
try:
    d = tencent._create_driver(profile=prof)
    print("成功:", type(d).__name__)
    print("打开腾讯文档...")
    d.get("https://docs.qq.com/sheet/DWGhDUEJsek1IcGNH")
    import time
    time.sleep(6)
    print("标题:", d.title)
    d.quit()
except Exception as e:
    print("失败:", type(e).__name__, str(e)[:500])
