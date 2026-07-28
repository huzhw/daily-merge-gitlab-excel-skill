# -*- coding: utf-8 -*-
"""daily-merge 公用工具模块"""

import os
import re

# 任务间分隔线，中文破折号 50 个顶满 D 列宽度
SEP = '—' * 50


def find_report_dir(base_path, year_str, month_str):
    """找或创建当月报告目录，兼容 07月 / 7月 两种格式。

    优先匹配带前导零的格式（日报-2026-07月），
    回退到不带前导零的格式（日报-2026-7月），
    都不存在时创建带前导零的目录。

    Args:
        base_path: 桌面根目录
        year_str:  四位年份字符串，如 "2026"
        month_str: 两位月份字符串（strftime("%m")），如 "07"

    Returns:
        报告目录的绝对路径
    """
    parent = os.path.join(base_path, f"报告-{year_str}年")

    # 优先匹配 07月 格式（当前规范）
    path_mm = os.path.join(parent, f"日报-{year_str}-{month_str}月")

    # 回退匹配 7月 格式（历史遗留）
    month_no_zero = str(int(month_str))
    path_m = os.path.join(parent, f"日报-{year_str}-{month_no_zero}月")

    if os.path.isdir(path_mm):
        return path_mm
    if os.path.isdir(path_m):
        return path_m

    # 都不存在 → 创建规范格式目录
    os.makedirs(path_mm, exist_ok=True)
    return path_mm


def to_chinese(n):
    """数字转中文数字，1→一, 10→十, 26→二十六，>99 回退数字"""
    digits = '零一二三四五六七八九'
    if n <= 0:
        return str(n)
    if n < 10:
        return digits[n]
    if n < 20:
        return '十' + (digits[n % 10] if n % 10 != 0 else '')
    if n < 100:
        tens = digits[n // 10]
        ones = digits[n % 10] if n % 10 != 0 else ''
        return tens + '十' + ones
    return str(n)


def format_desc(desc):
    """冒号后的行内编号子项 → 换行缩进，解决 D 列层级冲突。

    "任务描述：1.子项A 2.子项B" → "任务描述：\\n    1.子项A\\n    2.子项B"
    首个子项紧跟冒号无空格也能匹配；没有冒号时原样返回。
    """
    idx = max(desc.rfind('：'), desc.rfind(':'))
    if idx < 0:
        return desc
    before = desc[:idx + 1]
    after = desc[idx + 1:]
    # 首个子项紧跟冒号（无空格）
    after = re.sub(r'^(\d+[\.、)])', r'\n    \1', after)
    # 后续子项（有空格）
    after = re.sub(r'\s+(\d+[\.、)])', r'\n    \1', after)
    return before + after
