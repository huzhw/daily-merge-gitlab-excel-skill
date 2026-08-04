# -*- coding: utf-8 -*-
"""daily-merge 公用工具模块"""

import os
import re
from datetime import timedelta

# 任务间分隔线，中文破折号 50 个顶满 D 列宽度
SEP = '—' * 30

# G 列叠加时跳过：D 列含这些关键词 + F=G=J（当天创建当天完成）的临时任务
SKIP_KEYWORDS = ['开会', '部署', '沟通']

# chinesecalendar 未安装时的降级警告只打一次
_HOLIDAY_WARNED = False


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
    # md 表格内的换行标签 → Excel 单元格换行符（兼容 <br>、<br/>、<br />、</br>、大小写）
    desc = re.sub(r'<\s*/?\s*br\s*/?\s*>', '\n', desc, flags=re.I)
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


def is_temp_task(d_val, f_val, g_val, j_val):
    """D 列含跳过关键词 + F=G=J 同一天 → True，G 列叠加时跳过"""
    if not d_val:
        return False
    if not any(kw in str(d_val) for kw in SKIP_KEYWORDS):
        return False
    from datetime import datetime
    dates = []
    for v in (f_val, g_val, j_val):
        if v is None:
            return False
        if isinstance(v, datetime):
            dates.append(v.date())
        else:
            try:
                dates.append(datetime.strptime(str(v)[:10], '%Y-%m-%d').date())
            except:
                return False
    return dates[0] == dates[1] == dates[2]


def is_workday(d):
    """判断是否为工作日：跳过双休和法定节假日，调休上班日算工作日。

    G 列叠加日期时用。优先 chinesecalendar 库（含节假日+调休数据，
    按年维护）；库未安装时降级为仅跳过周六日，并打印一次警告。

    Args:
        d: date / datetime 对象

    Returns:
        True 表示工作日
    """
    global _HOLIDAY_WARNED
    try:
        import chinese_calendar as cc
        return cc.is_workday(d)
    except ImportError:
        if not _HOLIDAY_WARNED:
            print("[警告] chinesecalendar 未安装，G 列只跳双休、不跳节假日。"
                  "请执行: pip install chinesecalendar")
            _HOLIDAY_WARNED = True
        return d.weekday() < 5


def next_workday(d):
    """下一个工作日：从 d 的下一天起，跳过双休和法定节假日（含调休）"""
    d = d + timedelta(days=1)
    while not is_workday(d):
        d = d + timedelta(days=1)
    return d


def unmerge_notes_horizontal(ws):
    """删除所有 B~I 范围内的横向合并（填充说明区合并特征）。

    daily.py 必须在 load 后立即调用：openpyxl 的 insert_rows() 不移动已存在
    的合并单元格，残留横向合并在插入后可能落在新任务行上，而向合并区域内
    非锚点 cell 写值会丢失，导致新行 C~H 内容被吞。先清掉，保证后续
    insert_rows/写值全程无横向合并干扰，保存前再调用 rebuild_notes_merges 重建。

    Args:
        ws: openpyxl Worksheet 对象
    """
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row == mr.max_row and mr.min_col <= 2 and mr.max_col >= 9:
            ws.unmerge_cells(str(mr))


def rebuild_notes_merges(ws):
    """重建填充说明区的 B~I 横向合并，修复 insert_rows 错位。

    openpyxl 的 insert_rows() 只移动单元格值、不移动已存在的合并单元格，
    导致 daily.py 插行后说明区的横向合并（B~I）原地停留、与下移后的
    说明内容错位，横向合并可能压住新写入的任务行，把 D~I 内容吞掉。
    保存前调用本函数，按说明内容当前实际位置重建横向合并。

    Args:
        ws: openpyxl Worksheet 对象

    Returns:
        None
    """
    # 1. 删除所有 B~I 范围内的横向合并（说明区合并特征：单行 + 列跨 B~I 或更宽）
    unmerge_notes_horizontal(ws)

    # 2. 定位「填充说明」标题行
    notes_row = None
    for row in range(1, ws.max_row + 1):
        a = ws.cell(row=row, column=1).value
        if a and '填充说明' in str(a):
            notes_row = row
            break
    if notes_row is None:
        return

    # 3. 从标题行往下找最后一个非空行（标题在 A 列，说明条在 B 列）
    last_row = ws.max_row
    while last_row > notes_row:
        if ws.cell(row=last_row, column=1).value or ws.cell(row=last_row, column=2).value:
            break
        last_row -= 1

    # 4. 逐行重建 B~I 横向合并（跳过空白行）
    for r in range(notes_row, last_row + 1):
        if ws.cell(row=r, column=1).value or ws.cell(row=r, column=2).value:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
