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
    """层级缩进排版(与日报管家 sharedRules.js 同规则,改任一侧必须双侧同步)：
    <br>→换行;编号行(1、/1-1、/1-1-1、)按 dash 深度纯空格缩进,
    每级 6 空格(1、→6;1-1、→12;1-1-1、→18);标题/叙述行顶格;
    括号注释行缩进 6 放末尾。无装饰字符,弹窗与 Excel 同文本。
    """
    desc = re.sub(r'<\s*/?\s*br\s*/?\s*>', '\n', desc, flags=re.I)
    # 行内「冒号 + 编号子项」折行(编号后必须接空格,防时间/版本号误拆)
    desc = re.sub(r'[：:](?=\d+[\.、）)]\s)', '：\n', desc)
    out = []
    for raw in desc.split('\n'):
        line = raw.strip()
        if not line:
            continue
        m = re.match(r'^((?:\d+(?:-\d+)*))[、.)]?\s*(.*)$', line)
        if m and re.match(r'^\d+', line):
            # 编号统一补顿号(1 / 1. / 1-1 → 1、 / 1-1、),按父链深度缩进(每级 6 空格)
            level = m.group(1).count('-') + 1
            out.append(' ' * (level * 6) + m.group(1) + '、' + m.group(2))
        elif re.match(r'^[（(]', line):
            out.append('      ' + line)   # 括号注释(提交次数/耗时/行数等),缩进 6
        else:
            out.append(line)              # 标题/叙述行顶格无标记
    return '\n'.join(out)


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
        # 库未安装：整体降级为仅跳双休
        if not _HOLIDAY_WARNED:
            print("[警告] chinesecalendar 未安装，G 列仅跳双休、不跳节假日。"
                  "可执行: pip install chinesecalendar")
            _HOLIDAY_WARNED = True
        return d.weekday() < 5
    except NotImplementedError:
        # 库已安装，但查询年份超出官方数据范围（如 2027 尚未公布放假安排）。
        # 按架构师决定：不做估算，缺数据年份直接按双休算；
        # 官方数据发布后 pip install -U chinesecalendar 即自动恢复精确。
        if not _HOLIDAY_WARNED:
            print(f"[提示] {d.year} 年节假日数据未发布，该年日期仅按双休跳过"
                  "（官方数据发布后升级 chinesecalendar 自动生效）")
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
