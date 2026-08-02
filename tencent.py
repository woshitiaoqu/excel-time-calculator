# -*- coding: utf-8 -*-
"""抓取腾讯文档（公开分享链接）的数据。

原理：
- 腾讯文档公开分享链接可通过 dop-api/opendoc 接口匿名读取基础信息（工作簿名）。
- 但 opendoc 接口只返回部分单元格（websocket 才推送全量），因此用系统浏览器
  （Chrome / Edge / Firefox，无头模式）打开文档、依次点击每个工作表标签强制加载，
  再从页面全局对象 SpreadsheetApp.workbook 读取全部单元格，确保数据一个不少。
"""

import base64
import json
import os
import re
import time
import urllib.request
import urllib.parse
import zlib

OPENDOC_URL = "https://docs.qq.com/dop-api/opendoc"
SHEET_URL_RE = re.compile(r"docs\.qq\.com/(?:sheet|s)/([A-Za-z0-9_-]+)")
SHEET_ID_RE = re.compile(r"[A-Za-z0-9_-]{10,}")


def parse_doc_id(url):
    """从腾讯文档链接提取文档 ID。"""
    m = SHEET_URL_RE.search(url)
    if m:
        return m.group(1)
    m = SHEET_ID_RE.search(url)
    if m:
        return m.group(1)
    raise ValueError("无法识别腾讯文档链接：%s" % url)


def _open(url, referer):
    req = urllib.request.Request(url, headers={
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Referer": referer,
        "Origin": "https://docs.qq.com",
        "Accept": "*/*",
    })
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode("utf-8", "ignore")


def get_sheets(doc_id):
    """用 opendoc 接口快速获取工作表列表 [{"id":..., "name":...}, ...]。"""
    params = [
        ("u", ""), ("noEscape", "1"), ("startrow", "0"), ("endrow", "1000"),
        ("needSheetState", "1"), ("sliceStates", "1"), ("block_end_col", "100"),
        ("block_end_row", "1000"), ("block_start_col", "0"), ("block_start_row", "0"),
        ("id", doc_id), ("normal", "1"), ("outformat", "1"), ("wb", "1"),
        ("nowb", "0"), ("callback", "clientVarsCallback"), ("xsrf", ""), ("t", "0"),
    ]
    url = OPENDOC_URL + "?" + urllib.parse.urlencode(params)
    raw = _open(url, "https://docs.qq.com/sheet/%s" % doc_id)
    m = re.search(r"clientVarsCallback\((.*)\)\s*$", raw, re.S)
    if not m:
        raise ValueError("接口响应格式异常（可能链接失效或需登录）")
    data = json.loads(m.group(1))
    cv = data.get("clientVars", {})
    header = cv.get("collab_client_vars", {}).get("header", [])
    for h in header:
        d = h.get("d", [])
        if d and h.get("type") == "ms":
            return [{"id": s.get("id", ""), "name": s.get("name", "")}
                    for s in d]
    return []


# ---------------------------------------------------------------------------
# 浏览器全量获取
# ---------------------------------------------------------------------------

_CHROME_OPTIONS = [
    "--headless=new",
    "--disable-gpu",
    "--no-sandbox",
    "--disable-dev-shm-usage",
    "--window-size=1400,900",
    "--lang=zh-CN",
    "--no-proxy-server",
    "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
]

_EXTRACT_JS = """
var wb = window.SpreadsheetApp.workbook;
var sheets = wb.worksheetManager.sheetList;
var names = sheets.map(function(s) { return s._AnT; });
return JSON.stringify(names);
"""

_GET_CELLS_JS = """
var name = arguments[0];
var wb = window.SpreadsheetApp.workbook;
var s = wb.worksheetManager.sheetList.find(function(x) { return x._AnT === name; });
if (!s) return 'NO_SHEET';
var kk = s.cellDataGrid._kK;
var cols = {};
var brs = Object.keys(kk).map(Number).sort(function(a, b) { return a - b; });
for (var i = 0; i < brs.length; i++) {
  var bcs = Object.keys(kk[brs[i]]).map(Number).sort(function(a, b) { return a - b; });
  for (var j = 0; j < bcs.length; j++) {
    var arr = kk[brs[i]][bcs[j]]._Ao;
    if (!Array.isArray(arr)) continue;
    for (var r = 0; r < arr.length; r++) {
      var rowArr = arr[r];
      if (!Array.isArray(rowArr)) continue;
      for (var c = 0; c < rowArr.length; c++) {
        var cell = rowArr[c];
        if (!cell || typeof cell !== 'object') continue;
        var v = cell.value;
        if (v === undefined || v === null) continue;
        var letter = String.fromCharCode(65 + c);
        if (!cols[letter]) cols[letter] = [];
        cols[letter].push(String(v));
      }
    }
  }
}
return JSON.stringify(cols);
"""

_CLICK_TAB_JS = """
var name = arguments[0];
var els = document.querySelectorAll('.tab-bar-item');
for (var i = 0; i < els.length; i++) {
  if (els[i].textContent.trim() === name) { els[i].click(); return true; }
}
return false;
"""


def _sys_browser_paths():
    """自动检测系统已安装的浏览器，返回 [(kind, exe_path), ...]。

    通过注册表 + 常见安装目录自动发现 Chrome/Edge/Firefox 的位置，
    不依赖硬编码路径。
    """
    found = []

    def add(kind, exe):
        if exe and os.path.exists(exe) and exe not in [p for _, p in found]:
            found.append((kind, exe))

    # 1) 注册表 App Paths（最可靠）
    import winreg
    app_paths = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
    for exe_name in ("chrome.exe", "msedge.exe", "firefox.exe"):
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                app_paths + "\\" + exe_name) as k:
                exe = winreg.QueryValue(k, None)
                if exe:
                    kind = {"chrome.exe": "chrome",
                            "msedge.exe": "edge",
                            "firefox.exe": "firefox"}[exe_name]
                    add(kind, exe)
        except OSError:
            pass

    # 2) 常见安装目录
    pf = os.environ.get("ProgramFiles", "")
    pfx = os.environ.get("ProgramFiles(x86)", "")
    local = os.environ.get("LOCALAPPDATA", "")
    cands = [
        ("chrome", os.path.join(pf, "Google\\Chrome\\Application\\chrome.exe")),
        ("chrome", os.path.join(pfx, "Google\\Chrome\\Application\\chrome.exe")),
        ("chrome", os.path.join(local, "Google\\Chrome\\Application\\chrome.exe")),
        ("edge", os.path.join(pf, "Microsoft\\Edge\\Application\\msedge.exe")),
        ("edge", os.path.join(pfx, "Microsoft\\Edge\\Application\\msedge.exe")),
        ("firefox", os.path.join(pf, "Mozilla Firefox\\firefox.exe")),
        ("firefox", os.path.join(pfx, "Mozilla Firefox\\firefox.exe")),
    ]
    for kind, p in cands:
        add(kind, p)

    # 3) 按可执行文件名在 PATH 里找
    for kind, exe_name in (("chrome", "chrome.exe"), ("edge", "msedge.exe"),
                           ("firefox", "firefox.exe")):
        for d in os.environ.get("PATH", "").split(os.pathsep):
            if not d:
                continue
            p = os.path.join(d.strip('"'), exe_name)
            if os.path.exists(p):
                add(kind, p)

    # 4) 找不到时用默认浏览器
    if not found:
        add("edge", r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
        add("chrome", r"C:\Program Files\Google\Chrome\Application\chrome.exe")
        add("firefox", r"C:\Program Files\Mozilla Firefox\firefox.exe")

    return found


def _browser_exe_version(exe):
    """获取浏览器可执行文件的版本号（FileVersion）。"""
    try:
        info = os.popen('powershell -NoProfile -Command '
                        '"(Get-Item \'%s\').VersionInfo.FileVersion"'
                        % exe.replace("'", "''")).read().strip()
        if re.match(r"^\d+\.\d+", info):
            return info
    except Exception:
        pass
    return None


def _user_data_dirs():
    """自动查找各浏览器用户数据目录（含已登录 profile）。"""
    local = os.environ.get("LOCALAPPDATA", "")
    out = [
        ("chrome", os.path.join(local, "Google\\Chrome\\User Data")),
        ("edge", os.path.join(local, "Microsoft\\Edge\\User Data")),
    ]
    return [(k, p) for k, p in out if os.path.isdir(p)]


def _profile_cookies_db(profile_dir):
    """返回 (Cookies 数据库, 是否已加密) 路径。"""
    for sub in ("Network\\Cookies", "Cookies"):
        p = os.path.join(profile_dir, sub)
        if os.path.exists(p):
            return p
    return None


def _cookies_contain(profile_dir, hosts):
    """检查 profile 的 cookie 库是否包含给定域名（判断是否已登录）。"""
    db = _profile_cookies_db(profile_dir)
    if not db:
        return False
    try:
        import sqlite3
        con = sqlite3.connect("file:%s?mode=ro&immutable=1" % db, uri=True)
        try:
            cur = con.cursor()
            for host in hosts:
                try:
                    cur.execute(
                        "SELECT 1 FROM cookies WHERE host_key LIKE ? LIMIT 1",
                        ("%" + host + "%",))
                    if cur.fetchone():
                        return True
                except Exception:
                    continue
        finally:
            con.close()
    except Exception:
        return False
    return False


def _find_logged_in_profile(hosts):
    """自动扫描各浏览器 profile，返回第一个已登录目标域名的
    (kind, user_data_dir, profile_dir, exe_path)。"""
    browsers = dict(_sys_browser_paths())
    for kind, udd in _user_data_dirs():
        if not os.path.isdir(udd):
            continue
        # 扫描所有 profile 目录
        for entry in sorted(os.listdir(udd)):
            if entry in ("System Profile", "Guest Profile", "Default",
                         "Crashpad", "component_crx_cache",
                         "extensions_crx_cache"):
                continue
            prof = os.path.join(udd, entry)
            if not os.path.isdir(prof):
                continue
            if _cookies_contain(prof, hosts):
                return kind, udd, prof, browsers.get(kind)
    # Default profile 兜底
    for kind, udd in _user_data_dirs():
        prof = os.path.join(udd, "Default")
        if os.path.isdir(prof) and _cookies_contain(prof, hosts):
            return kind, udd, prof, browsers.get(kind)
    return None


def _chrome_driver():
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    opts = Options()
    for a in _CHROME_OPTIONS:
        opts.add_argument(a)
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    # 打包/缓存的驱动优先，失败时回退 selenium manager 自动匹配
    drv = _bundle_path("chromedriver.exe")
    if drv:
        try:
            return webdriver.Chrome(options=opts, service=Service(drv))
        except Exception:
            pass
    return webdriver.Chrome(options=opts)


def _firefox_driver():
    from selenium import webdriver
    from selenium.webdriver.firefox.options import Options
    from selenium.webdriver.firefox.service import Service
    opts = Options()
    opts.add_argument("--headless")
    drv = _bundle_path("geckodriver.exe")
    if drv:
        try:
            return webdriver.Firefox(options=opts, service=Service(drv))
        except Exception:
            pass
    return webdriver.Firefox(options=opts)


def _create_driver(profile=None):
    """自动选择可用浏览器 + 驱动 + profile，创建无头 driver。

    完全自动，不写死任何路径：
    - 驱动：优先用本机已缓存（selenium 缓存），没有则 selenium manager 自动下载
    - profile：若传入已登录 profile，复用其登录态；否则全新临时 profile
    - 浏览器内核：Edge 驱动可配 Chrome/Edge 的 profile，互相通用（同为 Chromium）
    """
    errors = []

    # 收集可用驱动
    drv_chrome = _bundle_path("chromedriver.exe")
    drv_edge = _bundle_path("msedgedriver.exe")
    drv_firefox = _bundle_path("geckodriver.exe")

    udd, prof = None, None
    if profile:
        _, udd, prof, _ = profile

    def build_chromium(use_edge_driver, udd=None, prof=None):
        from selenium import webdriver
        if use_edge_driver:
            from selenium.webdriver.edge.options import Options
            from selenium.webdriver.edge.service import Service
            opts = Options()
            driver_fn = webdriver.Edge
            drv = drv_edge
        else:
            from selenium.webdriver.chrome.options import Options
            from selenium.webdriver.chrome.service import Service
            opts = Options()
            driver_fn = webdriver.Chrome
            drv = drv_chrome
        for a in _CHROME_OPTIONS:
            opts.add_argument(a)
        opts.add_experimental_option("excludeSwitches", ["enable-logging"])
        if udd and prof:
            opts.add_argument("--user-data-dir=" + udd)
            opts.add_argument("--profile-directory=" + os.path.basename(prof))
        if drv:
            try:
                return driver_fn(options=opts, service=Service(drv))
            except Exception as e:
                errors.append("driver(%s): %s" % (os.path.basename(drv), str(e)[:80]))
        try:
            return driver_fn(options=opts)
        except Exception as e:
            errors.append("driver-auto: %s" % str(e)[:80])
        return None

    def build_firefox():
        from selenium import webdriver
        from selenium.webdriver.firefox.options import Options
        from selenium.webdriver.firefox.service import Service
        opts = Options()
        opts.add_argument("--headless")
        if drv_firefox:
            try:
                return webdriver.Firefox(options=opts, service=Service(drv_firefox))
            except Exception as e:
                errors.append("firefox-driver: %s" % str(e)[:80])
        try:
            return webdriver.Firefox(options=opts)
        except Exception as e:
            errors.append("firefox-auto: %s" % str(e)[:80])
        return None

    # 组合尝试顺序：edge驱动+profile -> chrome驱动+profile -> edge驱动裸 -> chrome驱动裸 -> firefox
    attempts = []
    if profile:
        attempts.append(lambda: build_chromium(True, udd, prof))   # edge drv + profile
        attempts.append(lambda: build_chromium(False, udd, prof))  # chrome drv + profile
    if drv_edge:
        attempts.append(lambda: build_chromium(True))
    if drv_chrome:
        attempts.append(lambda: build_chromium(False))
    attempts.append(lambda: build_chromium(True))    # edge auto
    attempts.append(lambda: build_chromium(False))   # chrome auto
    attempts.append(build_firefox)

    for fn in attempts:
        d = fn()
        if d is not None:
            return d

    raise RuntimeError(
        "未能启动无头浏览器，请确认电脑装有 Chrome、Edge 或 Firefox。\n"
        + "\n".join(errors[-5:]))


def _bundle_path(name):
    import sys as _sys
    base = getattr(_sys, "_MEIPASS", None)
    if base:
        p = os.path.join(base, name)
        if os.path.exists(p):
            return p
    return _cached_driver(name)


def _cached_driver(name):
    import glob as _glob
    home = os.path.expanduser("~")
    if name == "chromedriver":
        patterns = [
            os.path.join(home, ".cache\\selenium\\chromedriver\\win64\\*\\chromedriver.exe"),
            os.path.join(home, ".cache\\selenium\\chromedriver\\win32\\*\\chromedriver.exe"),
        ]
    elif name == "geckodriver":
        patterns = [os.path.join(home, ".cache\\selenium\\geckodriver\\win64\\*\\geckodriver.exe")]
    elif name == "msedgedriver":
        patterns = [
            os.path.join(home, ".cache\\selenium\\msedgedriver\\win64\\*\\msedgedriver.exe"),
            os.path.join(home, ".cache\\selenium\\msedgedriver\\win32\\*\\msedgedriver.exe"),
        ]
    else:
        patterns = []
    for pat in patterns:
        hits = sorted(_glob.glob(pat), key=os.path.getmtime, reverse=True)
        if hits:
            return hits[0]
    return None


def _pick_profile_for_url(url):
    """根据链接域名自动选择要复用的登录 profile。

    腾讯文档用 docs.qq.com 登录态；WPS 用 kdocs.cn / wps.cn 登录态。
    返回 profile 或 None（无需登录 / 未找到）。
    """
    hosts = []
    if "docs.qq.com" in url or "qq.com" in url:
        hosts = ["docs.qq.com", ".qq.com"]
    elif "kdocs.cn" in url or "wps.cn" in url or "kingsoft" in url:
        hosts = ["kdocs.cn", "wps.cn", "kingsoft"]
    if not hosts:
        return None
    return _find_logged_in_profile(hosts)


def fetch_workbook_with_browser(url, wait_sheet_ready=60, wait_after_click=4,
                                progress=None):
    """用系统无头浏览器打开在线文档，返回所有工作簿的所有列数据。

    自动识别平台（腾讯文档 / WPS 金山文档）并选择对应的读取方式：
    1. 检测系统浏览器（Chrome/Edge/Firefox）路径与版本，自动匹配 driver
    2. 检测浏览器是否已登录目标站点，复用登录态免登录
    3. 打开文档 → 遍历每个工作表 → 读取全部列单元格

    返回 [{"name": ..., "columns": {...}}, ...]。
    progress 回调：progress(msg)，用于 UI 显示"正在获取数据中"。
    """
    if progress:
        progress("正在检测浏览器…")
    profile = _pick_profile_for_url(url)
    if profile:
        if progress:
            progress("检测到已登录的浏览器，正在启动…")
    else:
        if progress:
            progress("正在启动浏览器…")
    driver = _create_driver(profile=profile)
    try:
        if progress:
            progress("正在打开在线文档…")
        driver.set_page_load_timeout(60)
        driver.get(url)
        time.sleep(5)

        if is_wps_url(url):
            return _fetch_wps(driver, wait_sheet_ready, progress)

        if progress:
            progress("正在等待文档加载完成…")
        ready = False
        for _ in range(wait_sheet_ready):
            try:
                r = driver.execute_script(
                    "return !!(window.SpreadsheetApp && "
                    "window.SpreadsheetApp.workbook && "
                    "window.SpreadsheetApp.workbook.worksheetManager)")
                if r:
                    ready = True
                    break
            except Exception:
                pass
            time.sleep(1)
        if not ready:
            raise RuntimeError("在线文档加载超时，请检查链接是否可访问")

        names = json.loads(driver.execute_script(_EXTRACT_JS))
        if progress:
            progress("共找到 %d 个工作簿，正在读取数据…" % len(names))

        result = []
        for name in names:
            driver.execute_script(_CLICK_TAB_JS, name)
            time.sleep(wait_after_click)
            raw = driver.execute_script(_GET_CELLS_JS, name)
            columns = {} if raw == "NO_SHEET" else json.loads(raw)
            result.append({"name": name, "columns": columns})
            total = sum(len(v) for v in columns.values())
            if progress:
                progress("工作簿「%s」已读取 %d 列 %d 条" % (
                    name, len(columns), total))
        return result
    finally:
        driver.quit()


def is_wps_url(url):
    return ("kdocs.cn" in url or "wps.cn" in url or "kingsoft" in url)


# ---------------------------------------------------------------------------
# WPS 读取逻辑（使用 WPSOpenApi）
# ---------------------------------------------------------------------------

_WPS_GET_SHEETS_JS = """
var els = document.querySelectorAll('.sheet-name');
var names = [];
for (var i = 0; i < els.length; i++) {
  var t = (els[i].textContent || '').trim();
  if (t) names.push(t);
}
return JSON.stringify(names);
"""

_WPS_ACTIVATE_JS = """
var idx = arguments[0];
var cb = arguments[arguments.length - 1];
var app = window.WPSOpenApi.Application;
try {
  var sh = app.Sheets.Item(idx);
  var p = sh.Activate();
  function done(){ cb(JSON.stringify({ok: true})); }
  if (p && p.then) p.then(done, function(e){ cb(JSON.stringify({err: String(e)})); });
  else done();
} catch(e) { cb(JSON.stringify({err: e.message})); }
"""

_WPS_READ_COL_JS = """
var letter = arguments[0];
var cb = arguments[arguments.length - 1];
var app = window.WPSOpenApi.Application;
var vals = [];
var row = 1;
var MAX_ROW = 500;
var dead = false;
function finish() { if (!dead) { dead = true; cb(JSON.stringify(vals)); } }
function readNext() {
  if (dead || row > MAX_ROW) { finish(); return; }
  var cell, t;
  try { cell = app.Range(letter + row); } catch(e) { finish(); return; }
  try { t = cell.Text; } catch(e) { finish(); return; }
  // 每个 promise 加 3 秒兜底，避免卡死
  var guard = setTimeout(function(){ finish(); }, 3000);
  function got(v) {
    clearTimeout(guard);
    if (dead) return;
    var s = v === undefined || v === null ? '' : String(v);
    if (s.indexOf('function') === 0) { finish(); return; }
    if (s === '') {
      if (vals.length && row > 2) { finish(); return; }
      row++; readNext(); return;
    }
    vals.push(s);
    row++; readNext();
  }
  function onErr() { clearTimeout(guard); finish(); }
  if (t && t.then) { t.then(got, onErr); }
  else { got(t); }
}
readNext();
"""

_WPS_GET_COLUMNS_JS = """
var cb = arguments[arguments.length - 1];
var app = window.WPSOpenApi.Application;
var cols = {};
var MAX_COL = 26;
var MAX_ROW = 500;
var c = 0;
function nextCol() {
  if (c >= MAX_COL) { cb(JSON.stringify(cols)); return; }
  var letter = String.fromCharCode(65 + c);
  var vals = [];
  var row = 1;
  function readRow() {
    if (row > MAX_ROW) { if (vals.length) cols[letter] = vals; c++; nextCol(); return; }
    var cell, t;
    try { cell = app.Range(letter + row); } catch(e) { if (vals.length) cols[letter] = vals; c++; nextCol(); return; }
    try { t = cell.Text; } catch(e) { if (vals.length) cols[letter] = vals; c++; nextCol(); return; }
    function got(v) {
      var s = v === undefined || v === null ? '' : String(v);
      if (s.indexOf('function') === 0) { if (vals.length) cols[letter] = vals; c++; nextCol(); return; }
      if (s === '') {
        if (vals.length && row > 2) { if (vals.length) cols[letter] = vals; c++; nextCol(); return; }
        row++; readRow(); return;
      }
      vals.push(s);
      row++; readRow();
    }
    if (t && t.then) { t.then(got, function(){ if (vals.length) cols[letter] = vals; c++; nextCol(); }); }
    else { got(t); }
  }
  readRow();
}
nextCol();
"""


def _fetch_wps(driver, wait_sheet_ready, progress):
    """读取 WPS 文档：遍历所有工作表，读取所有列。"""
    if progress:
        progress("正在等待 WPS 文档加载完成…")

    ready = False
    for _ in range(wait_sheet_ready):
        try:
            r = driver.execute_script(
                "return !!(window.WPSOpenApi && "
                "window.WPSOpenApi.Application)")
            if r:
                ready = True
                break
        except Exception:
            pass
        time.sleep(1)
    if not ready:
        raise RuntimeError("WPS 文档加载超时，请检查链接是否可访问")

    time.sleep(3)  # 等 API 完全就绪

    names = json.loads(driver.execute_script(_WPS_GET_SHEETS_JS))
    if not names:
        raise RuntimeError("未能识别 WPS 文档的工作表")
    if progress:
        progress("共找到 %d 个工作簿，正在读取数据…" % len(names))

    result = []
    for idx, name in enumerate(names, 1):
        if idx > 1:
            driver.execute_async_script(_WPS_ACTIVATE_JS, idx)
            time.sleep(2)
        columns = {}
        for c in range(26):
            letter = chr(65 + c)
            raw = driver.execute_async_script(_WPS_READ_COL_JS, letter)
            vals = json.loads(raw)
            if not vals:
                if c > 0:
                    break  # 空列后不再继续
                continue
            columns[letter] = vals
        result.append({"name": name, "columns": columns})
        total = sum(len(v) for v in columns.values())
        if progress:
            progress("工作簿「%s」已读取 %d 列 %d 条" % (
                name, len(columns), total))
    return result
