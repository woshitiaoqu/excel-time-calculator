# -*- coding: utf-8 -*-
"""求和与格式化。"""


def total_from_values(values):
    total = 0
    parsed = 0
    skipped = 0
    for v, time_cell in values:
        sec = parse_seconds(v, time_cell)
        if sec is None:
            skipped += 1
        else:
            total += sec
            parsed += 1
    return total, parsed, skipped


def parse_seconds(value, time_cell):
    from reader import parse_seconds as _p
    return _p(value, time_cell)


import math


def _fmt_hours(hours):
    # 截断两位小数，不进位（如 1.0344 -> 1.03，1.999 -> 1.99）
    v = math.floor(hours * 100) / 100
    s = "%.2f" % v
    s = s.rstrip("0").rstrip(".")
    return s if s else "0"


def build_result_lines(total, parsed, skipped):
    hours = total / 3600.0
    minutes = total / 60.0
    h = total // 3600
    m = (total % 3600) // 60
    sec = total % 60
    lines = [
        "共统计：%d 行（跳过 %d 行）" % (parsed, skipped),
        "总小时：%sh" % _fmt_hours(hours),
        "总分钟：%d 分钟" % int(minutes),
        "总秒数：%d 秒" % total,
        "",
        "即 %d 小时 %d 分 %d 秒" % (h, m, sec),
    ]
    return lines


def build_total_text(total):
    h = total // 3600
    m = (total % 3600) // 60
    sec = total % 60
    hours = total / 3600.0
    return "总计：%sh（%d 小时 %d 分 %d 秒）" % (_fmt_hours(hours), h, m, sec)
