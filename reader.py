# -*- coding: utf-8 -*-
"""读取 Excel / CSV 文件，并把每一格的内容解析成"秒"。"""

import os
import re
import csv
import datetime

EXT_KIND = {
    ".xlsx": "xlsx",
    ".xlsm": "xlsx",
    ".xls": "xls",
    ".csv": "csv",
}


class CsvData:
    def __init__(self, rows, encoding):
        self.rows = rows
        self.encoding = encoding


def _detect_encoding(path):
    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "utf-8", "gbk"):
        try:
            raw.decode(enc)
            return enc
        except UnicodeDecodeError:
            continue
    return "utf-8"


def load_file(path):
    ext = os.path.splitext(path)[1].lower()
    kind = EXT_KIND.get(ext)
    if kind is None:
        raise ValueError("不支持的格式：%s（仅支持 .xlsx / .xls / .csv）" % ext)

    if kind == "xlsx":
        import openpyxl
        wb = openpyxl.load_workbook(path)
        return wb, wb.sheetnames, kind

    if kind == "xls":
        import xlrd
        try:
            wb = xlrd.open_workbook(path, formatting_info=True)
        except Exception:
            wb = xlrd.open_workbook(path)
        return wb, wb.sheet_names(), kind

    # csv
    encoding = _detect_encoding(path)
    with open(path, "r", encoding=encoding, newline="") as f:
        rows = list(csv.reader(f))
    return CsvData(rows, encoding), ["数据"], kind


def _is_time_cell(value, fmt):
    if isinstance(value, (datetime.time, datetime.datetime, datetime.timedelta)):
        return True
    f = (fmt or "").lower()
    return ("h" in f) or ("[m]" in f) or ("[s]" in f)


def read_all_columns(handle, kind, sheet_name):
    """读取工作表的所有列，返回 {"A": [...], "B": [...], ...}。

    只返回有非空值的列，每列是按行排列的单元格值列表。
    """
    if kind == "xlsx":
        ws = handle[sheet_name]
        cols = {}
        for c in range(1, ws.max_column + 1):
            letter = get_column_letter(c)
            vals = []
            for r in range(1, ws.max_row + 1):
                cell = ws.cell(r, c)
                v = cell.value
                if v is None:
                    continue
                if isinstance(v, str) and not v.strip():
                    continue
                vals.append((v, _is_time_cell(v, cell.number_format)))
            if vals:
                cols[letter] = vals
        return cols

    if kind == "xls":
        import xlrd
        sh = handle.sheet_by_name(sheet_name)
        cols = {}
        for c in range(sh.ncols):
            letter = get_column_letter(c + 1)
            vals = []
            for r in range(sh.nrows):
                cell = sh.cell(r, c)
                v = cell.value
                if v is None or v == "":
                    continue
                vals.append((v, cell.ctype == xlrd.XL_CELL_DATE))
            if vals:
                cols[letter] = vals
        return cols

    # csv：只有一列
    return {"A": read_column_a(handle, kind, sheet_name)}


def get_column_letter(n):
    """1 -> A, 2 -> B, ... 27 -> AA"""
    letter = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letter = chr(65 + rem) + letter
    return letter


_FULLWIDTH = {ord("："): ":", ord("．"): ".", ord("，"): ",", ord("　"): " "}
_FULLWIDTH.update({ord("０") + i: str(i) for i in range(10)})


def _normalize_symbols(s):
    """全角符号/数字转半角，避免中英文混用导致解析失败。"""
    return s.translate(_FULLWIDTH)


def _parse_time_text(s):
    """把多种写法的 分:秒 文本解析成秒；解析不了返回 None。

    支持：2:27 / 2：27 / 2.27 / 2．27 / 2,27 / 2分27秒 / 2m27s / 2'27"
    也支持三段的 时:分:秒 作为兜底（无小时场景一般不出现）。
    """
    # 文字写法：2分27秒 / 2m27s / 2'27"，必须含标记字符才按此解析
    #（否则纯整数 90 会被误拆成 9分0秒）
    if re.search(r"[分秒ms\u2032\u2033'\"`]", s):
        m = re.match(r"^(\d+)\s*[分m\u2032'\u0060]?\s*(\d{1,2})\s*[秒s\u2033\"]?$", s)
        if m:
            minutes, seconds = int(m.group(1)), int(m.group(2))
            if seconds < 60:
                return minutes * 60 + seconds

    # 数字分隔写法：2:27 / 2.27 / 2,27
    parts = re.split(r"[:：.．,，]", s)
    if len(parts) >= 2:
        try:
            nums = [int(float(p)) for p in parts]
        except ValueError:
            return None
        if len(nums) == 2:
            m, sec = nums
            if sec < 60:
                return m * 60 + sec
            return None
        if len(nums) == 3:
            h, m, sec = nums
            if sec < 60 and m < 60:
                return h * 3600 + m * 60 + sec
            return None
    return None


def parse_seconds(value, time_cell=False):
    if value is None:
        return None

    if isinstance(value, datetime.timedelta):
        return int(round(value.total_seconds()))

    if isinstance(value, (datetime.time, datetime.datetime)):
        # 无小时场景：Excel 时间 3:40 显示值按 3分40秒 处理（时->分，分->秒）
        return value.hour * 60 + value.minute

    if isinstance(value, (int, float)):
        if time_cell:
            # Excel 时间序列值（1 天 = 1.0），按 MM:SS 解释：serial*1440 = 秒
            return int(round(value * 1440))
        if value == int(value):
            return int(value)  # 纯整数按秒
        # 小数按 分.秒 解释（如 0.27 = 0分27秒 = 27 秒）
        sec = _parse_time_text(str(value))
        if sec is not None:
            return sec
        return int(round(value))

    if isinstance(value, str):
        s = _normalize_symbols(value).strip()
        if not s:
            return None
        sec = _parse_time_text(s)
        if sec is not None:
            return sec
        try:
            return int(round(float(s)))  # 纯数字按秒
        except ValueError:
            return None

    return None


def write_total(handle, kind, sheet_name, text, path):
    if kind == "xlsx":
        from openpyxl.styles import Font
        ws = handle[sheet_name]
        last = 0
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if v is not None and str(v).strip() != "":
                last = r
        cell = ws.cell(last + 1, 1)
        cell.value = text
        cell.font = Font(bold=True)
        handle.save(path)
        return True

    if kind == "csv":
        rows = handle.rows
        rows.append([text])
        with open(path, "w", encoding=handle.encoding, newline="") as f:
            csv.writer(f).writerows(rows)
        return True

    return False
